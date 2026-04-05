// ============================================================
// HOME DOCUMENT AUTO-ORGANIZER
// Google Apps Script + Gemini AI Classification
// ============================================================
// SETUP:
// 1. Go to https://script.google.com and create a new project
// 2. Paste this entire script into Code.gs
// 3. Enable the Advanced Drive Service:
//    - In the Apps Script editor, click "Services" (+) on the left
//    - Find "Drive API" and click "Add" (Make sure it is v3)
// 4. Get your ROOT_FOLDER_ID: in Google Drive, open
//    My Drive > Documents > My_Docs > Wickman_Glade_Hill_House_Docs,
//    copy the folder ID from the URL, and paste it into getConfig()
// 5. Run setupFolderStructure() once to create subfolders inside the existing root
// 6. Run setApiKey() to store your Gemini API key
// 7. Run setupTrigger() once to enable auto-processing
// ============================================================

// ── CONFIGURATION ──────────────────────────────────────────

/**
 * Returns the configuration object.
 * Paste your Wickman_Glade_Hill_House_Docs folder ID into ROOT_FOLDER_ID.
 * (Get it from the URL: drive.google.com/drive/folders/THIS_PART)
 */
function getConfig() {
  return {
    // My Drive > Documents > My_Docs > Wickman_Glade_Hill_House_Docs
    ROOT_FOLDER_ID: '1f7l7o0WBQ6etkA98QR7FWm-eqvNuAo82',

    // How often the trigger checks for new files (in minutes)
    TRIGGER_INTERVAL_MINUTES: 10,

    // Max characters of document text to send to Gemini
    MAX_TEXT_LENGTH: 50000,

    // Suffix appended to copied filenames (before the extension)
    COPY_SUFFIX: '-cp',

    // Max execution time in ms (5 min safety margin under 6 min Apps Script limit)
    MAX_EXECUTION_MS: 5 * 60 * 1000,

    // Delay between Gemini API calls in ms (rate limiting)
    // Set to 4000 to respect the 15 Requests Per Minute limit on the free tier
    GEMINI_DELAY_MS: 4000,

    // Max file size in bytes to attempt OCR/conversion (skip larger files to Unsorted)
    MAX_FILE_SIZE_BYTES: 25 * 1024 * 1024, // 25 MB

    // Max entries in processed file tracking before pruning oldest
    // Capped at 1500 to prevent exceeding the 500KB total UserProperties quota
    MAX_PROCESSED_ENTRIES: 1500,
  };
}

/**
 * Folder structure definition.
 * Flat — no sub-folders.
 */
function getFolderStructure() {
  return [
    '01 - Purchase & Title',
    '02 - Mortgage & Finance',
    '03 - Insurance',
    '04 - Tax & Assessment',
    '05 - HOA',
    '06 - Maintenance & Improvements',
    '07 - Warranties',
    '08 - Utilities & Services',
  ];
}

/**
 * Classification categories sent to Gemini.
 * Maps AI response labels → folder names in the structure.
 */
function getCategoryMap() {
  return {
    'purchase_title': '01 - Purchase & Title',
    'mortgage_finance': '02 - Mortgage & Finance',
    'insurance': '03 - Insurance',
    'tax_assessment': '04 - Tax & Assessment',
    'hoa': '05 - HOA',
    'maintenance_improvements': '06 - Maintenance & Improvements',
    'warranty': '07 - Warranties',
    'utilities_services': '08 - Utilities & Services',
  };
}


// ── PROCESSED FILE TRACKING ────────────────────────────────

// UserProperties + chunking: each property value is capped at 9 KB,
// so we split the JSON across multiple keys.
const CHUNK_PREFIX = 'PROCESSED_CHUNK_';
const CHUNK_COUNT_KEY = 'PROCESSED_CHUNK_COUNT';
const CHUNK_SIZE_CHARS = 8000; // character count; safely under the 9 KB per-value limit for ASCII-dominant content

/**
 * Get the set of file IDs that have already been processed.
 * Reassembles the JSON from chunked UserProperties.
 */
function getProcessedFileIds() {
  const userProps = PropertiesService.getUserProperties();
  const countRaw = userProps.getProperty(CHUNK_COUNT_KEY);
  const count = countRaw ? parseInt(countRaw, 10) : 0;

  if (count === 0) return {};

  let json = '';
  for (let i = 0; i < count; i++) {
    json += (userProps.getProperty(CHUNK_PREFIX + i) || '');
  }

  try {
    return JSON.parse(json);
  } catch (e) {
    return {};
  }
}

/**
 * Save the entire processed file map back to UserProperties in chunks.
 * Prunes oldest entries if the map exceeds MAX_PROCESSED_ENTRIES.
 */
function saveProcessedFileIds(processed) {
  const config = getConfig();
  const entries = Object.entries(processed);

  if (entries.length > config.MAX_PROCESSED_ENTRIES) {
    // Sort by processedAt ascending, drop oldest
    entries.sort((a, b) => (a[1].processedAt || '').localeCompare(b[1].processedAt || ''));
    const keep = entries.slice(entries.length - config.MAX_PROCESSED_ENTRIES);
    processed = Object.fromEntries(keep);
    Logger.log('[CLEANUP] Pruned processed tracking to ' + config.MAX_PROCESSED_ENTRIES + ' entries');
  }

  const json = JSON.stringify(processed);
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK_SIZE_CHARS) {
    chunks.push(json.substring(i, i + CHUNK_SIZE_CHARS));
  }

  const userProps = PropertiesService.getUserProperties();

  // Delete any old chunks that exceed the new count
  const oldCountRaw = userProps.getProperty(CHUNK_COUNT_KEY);
  const oldCount = oldCountRaw ? parseInt(oldCountRaw, 10) : 0;
  for (let i = chunks.length; i < oldCount; i++) {
    userProps.deleteProperty(CHUNK_PREFIX + i);
  }

  // Write new chunks and update count
  for (let i = 0; i < chunks.length; i++) {
    userProps.setProperty(CHUNK_PREFIX + i, chunks[i]);
  }
  userProps.setProperty(CHUNK_COUNT_KEY, String(chunks.length));
}

/**
 * Run this to clear the processed file tracking.
 * Useful if you want to reprocess everything from scratch.
 */
function resetProcessedTracking() {
  const userProps = PropertiesService.getUserProperties();
  const countRaw = userProps.getProperty(CHUNK_COUNT_KEY);
  const count = countRaw ? parseInt(countRaw, 10) : 0;

  for (let i = 0; i < count; i++) {
    userProps.deleteProperty(CHUNK_PREFIX + i);
  }
  userProps.deleteProperty(CHUNK_COUNT_KEY);
  Logger.log('[SUCCESS] Processed file tracking reset. All files will be reprocessed on next run.');
}


// ── SETUP FUNCTIONS ────────────────────────────────────────

/**
 * Run this ONCE to create subfolders inside your existing
 * Wickman_Glade_Hill_House_Docs folder.
 * Requires ROOT_FOLDER_ID to be set in getConfig() first.
 * Skips any folders that already exist.
 */
function setupFolderStructure() {
  const config = getConfig();
  const categories = getFolderStructure();

  if (!config.ROOT_FOLDER_ID) {
    Logger.log('[ERROR] ROOT_FOLDER_ID not set in getConfig().');
    Logger.log('   Open Wickman_Glade_Hill_House_Docs in Drive,');
    Logger.log('   copy the folder ID from the URL, and paste it in.');
    return;
  }

  const root = DriveApp.getFolderById(config.ROOT_FOLDER_ID);
  Logger.log('[SUCCESS] Using existing root: ' + root.getName() + ' (' + root.getId() + ')');

  // Helper: create folder only if it doesn't already exist
  function ensureFolder(parent, name) {
    const existing = parent.getFoldersByName(name);
    if (existing.hasNext()) {
      Logger.log('[SKIP] Already exists: ' + name);
      return existing.next();
    }
    const created = parent.createFolder(name);
    Logger.log('[SUCCESS] Created: ' + name);
    return created;
  }

  // Create system folders
  ensureFolder(root, 'Originals');
  ensureFolder(root, 'Inbox');
  ensureFolder(root, 'Unsorted');

  // Create category folders (flat, no sub-folders)
  for (const category of categories) {
    ensureFolder(root, category);
  }

  Logger.log('');
  Logger.log('[INFO] Folder structure complete!');
  Logger.log('');
  Logger.log('NEXT STEPS:');
  Logger.log('1. Set your Gemini API key: Run setApiKey()');
  Logger.log('2. Run setupTrigger() to start auto-processing');
}

/**
 * Run this to store your Gemini API key securely.
 * Get your key at: https://aistudio.google.com/app/apikey
 */
function setApiKey() {
  // [WARNING] Replace the string below with your actual API key, run once, then delete the key from this code
  const apiKey = 'PASTE_YOUR_GEMINI_API_KEY_HERE';

  if (apiKey === 'PASTE_YOUR_GEMINI_API_KEY_HERE') {
    Logger.log('[ERROR] Replace the placeholder with your actual Gemini API key first');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
  Logger.log('[SUCCESS] API key stored securely in Script Properties');
  Logger.log('[WARNING] Now delete the key from this function to keep it safe');
}

/**
 * Run this ONCE to set up the time-driven trigger.
 */
function setupTrigger() {
  // Remove any existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processOriginals') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const config = getConfig();
  const validIntervals = [1, 5, 10, 15, 30];
  if (!validIntervals.includes(config.TRIGGER_INTERVAL_MINUTES)) {
    Logger.log('[ERROR] TRIGGER_INTERVAL_MINUTES must be one of: ' + validIntervals.join(', ') + '. Got: ' + config.TRIGGER_INTERVAL_MINUTES);
    return;
  }

  ScriptApp.newTrigger('processOriginals')
    .timeBased()
    .everyMinutes(config.TRIGGER_INTERVAL_MINUTES)
    .create();

  Logger.log('[SUCCESS] Trigger set: processOriginals will run every ' + config.TRIGGER_INTERVAL_MINUTES + ' minutes');
}

/**
 * Run this to remove the auto-processing trigger.
 */
function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processOriginals') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  Logger.log('[SUCCESS] Removed ' + removed + ' trigger(s)');
}


// ── MAIN PROCESSING ────────────────────────────────────────

/**
 * Main function that runs on the trigger.
 *
 * Flow:
 * 1. Scan Originals for new (unprocessed) files
 * 2. Copy each new file into Inbox (with -cp suffix)
 * 3. Classify the Inbox copy using Gemini
 * 4. Move the Inbox copy into the correct category folder (01-08) or Unsorted
 * 5. Inbox empties out when processing is complete
 *
 * Originals are NEVER moved or modified.
 */
function processOriginals() {
  // Prevent concurrent trigger runs from duplicating work
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('[SKIP] Another run is in progress, skipping.');
    return;
  }

  try {
    _processOriginalsImpl();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Internal implementation — called by processOriginals() under lock.
 */
function _processOriginalsImpl() {
  const config = getConfig();
  const startTime = Date.now();

  if (!config.ROOT_FOLDER_ID) {
    Logger.log('[ERROR] ROOT_FOLDER_ID not set. Run setupFolderStructure() first.');
    return;
  }

  const rootFolder = DriveApp.getFolderById(config.ROOT_FOLDER_ID);
  const originalsFolder = getSubFolder(rootFolder, 'Originals');
  const inboxFolder = getSubFolder(rootFolder, 'Inbox');
  const unsortedFolder = getSubFolder(rootFolder, 'Unsorted');

  if (!originalsFolder) {
    Logger.log('[ERROR] Originals folder not found');
    return;
  }

  if (!inboxFolder) {
    Logger.log('[ERROR] Inbox folder not found');
    return;
  }

  if (!unsortedFolder) {
    Logger.log('[ERROR] Unsorted folder not found');
    return;
  }

  // Load processed tracking once into memory
  const processed = getProcessedFileIds();

  const files = originalsFolder.getFiles();
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  let unsortedCount = 0;
  let timedOut = false;

  try {
    while (files.hasNext()) {
      // Guard against Apps Script 6-minute execution limit
      if (Date.now() - startTime > config.MAX_EXECUTION_MS) {
        Logger.log('[TIMEOUT] Approaching time limit, stopping. Remaining files will process next run.');
        timedOut = true;
        break;
      }

      const file = files.next();
      const fileId = file.getId();

      // Skip files we've already processed
      if (Object.prototype.hasOwnProperty.call(processed, fileId)) {
        skippedCount++;
        continue;
      }

      Logger.log('[INFO] Processing: ' + file.getName());

      let inboxCopy = null;

      try {
        // Step 1: Copy file to Inbox with suffix
        inboxCopy = copyFileToFolder(file, inboxFolder, config.COPY_SUFFIX);
        Logger.log('   [COPIED] Copied to Inbox: ' + inboxCopy.getName());

        // Rate limit Gemini calls
        if (processedCount > 0 || unsortedCount > 0) {
          Utilities.sleep(config.GEMINI_DELAY_MS);
        }

        // Step 2: Classify and move the Inbox copy
        const result = classifyAndMoveFile(inboxCopy, rootFolder, unsortedFolder, config);

        if (result.sorted) {
          Logger.log('   [SUCCESS] Moved to: ' + result.destination);
          processed[fileId] = {
            classification: result.classification,
            processedAt: new Date().toISOString(),
          };
          processedCount++;
        } else {
          Logger.log('   [UNSORTED] Unsorted: ' + result.reason);
          processed[fileId] = {
            classification: 'unsorted',
            processedAt: new Date().toISOString(),
          };
          unsortedCount++;
        }
      } catch (error) {
        Logger.log('   [ERROR] Error: ' + error.message);
        errorCount++;
        // Clean up orphaned inbox copy so duplicates don't accumulate
        if (inboxCopy) {
          try {
            inboxCopy.setTrashed(true); // Optimized to skip DriveApp redundant lookup
            Logger.log('   [CLEANUP] Cleaned up failed inbox copy');
          } catch (cleanupError) {
            Logger.log('   [WARNING] Could not clean up inbox copy: ' + cleanupError.message);
          }
        }
        // Don't mark as processed on error — will retry next run
      }
    }
  } finally {
    // Always write processed map back, even if an unexpected exception occurs mid-loop
    saveProcessedFileIds(processed);
  }

  if (processedCount > 0 || unsortedCount > 0 || errorCount > 0) {
    Logger.log('');
    Logger.log('[SUMMARY] ' + processedCount + ' sorted, ' + unsortedCount + ' unsorted, ' + skippedCount + ' already processed, ' + errorCount + ' errors' + (timedOut ? ' (timed out, will continue next run)' : ''));
  }
}

/**
 * Classify a single file using Gemini and MOVE it to the correct folder.
 * The file passed in is the Inbox copy — it gets moved out of Inbox.
 */
function classifyAndMoveFile(file, rootFolder, unsortedFolder, config) {
  const fileName = file.getName();

  // Extract text from the file
  const text = extractText(file, config);

  if (!text || text.trim().length < 20) {
    moveFileToFolder(file, unsortedFolder);
    return { sorted: false, classification: 'unsorted', reason: 'Could not extract enough text for classification' };
  }

  // Classify using Gemini
  const classification = classifyWithGemini(text, fileName, config);

  if (!classification || classification === 'unknown') {
    moveFileToFolder(file, unsortedFolder);
    return { sorted: false, classification: 'unknown', reason: 'Gemini could not confidently classify this document' };
  }

  // Look up the destination folder name
  const categoryMap = getCategoryMap();
  const folderName = categoryMap[classification];

  if (!folderName) {
    moveFileToFolder(file, unsortedFolder);
    return { sorted: false, classification: classification, reason: 'Classification "' + classification + '" not mapped to a folder' };
  }

  // Find the destination folder
  const destFolder = getSubFolder(rootFolder, folderName);

  if (!destFolder) {
    moveFileToFolder(file, unsortedFolder);
    return { sorted: false, classification: classification, reason: 'Destination folder not found: ' + folderName };
  }

  // Move the file from Inbox to the category folder
  moveFileToFolder(file, destFolder);
  return { sorted: true, classification: classification, destination: folderName };
}


// ── TEXT EXTRACTION ────────────────────────────────────────

/**
 * Extract text content from a file.
 * Supports PDFs (via Drive OCR), Google Docs, text files, DOCX, and images.
 */
function extractText(file, config) {
  const mimeType = file.getMimeType();
  const fileName = file.getName();
  let text = '';

  // Skip files too large for OCR/conversion (avoids timeouts and quota issues)
  const fileSize = file.getSize();
  if (fileSize > config.MAX_FILE_SIZE_BYTES) {
    Logger.log('   [SKIP] File too large for extraction (' + Math.round(fileSize / 1024 / 1024) + ' MB), using filename only');
    return 'Filename: ' + fileName + '\n[File too large for text extraction: ' + Math.round(fileSize / 1024 / 1024) + ' MB]';
  }

  try {
    if (mimeType === 'application/pdf') {
      text = extractViaConversion(file, { ocrLanguage: 'en' });
    } else if (mimeType === 'application/vnd.google-apps.document') {
      const doc = DocumentApp.openById(file.getId());
      text = doc.getBody().getText();
    } else if (mimeType === 'text/plain' || mimeType === 'text/csv') {
      text = file.getBlob().getDataAsString();
    } else if (
      mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      mimeType === 'application/msword' ||
      mimeType === 'application/rtf' ||
      mimeType === 'text/rtf' ||
      mimeType === 'text/html'
    ) {
      text = extractViaConversion(file);
    } else if (mimeType.startsWith('image/')) {
      text = extractViaConversion(file, { ocrLanguage: 'en' });
    } else {
      text = '';
    }
  } catch (e) {
    Logger.log('   [WARNING] Text extraction fallback for ' + mimeType + ': ' + e.message);
    text = '';
  }

  // Prepend filename — often contains useful classification hints
  text = 'Filename: ' + fileName + '\n\n' + text;

  // Truncate to max length
  if (text.length > config.MAX_TEXT_LENGTH) {
    text = text.substring(0, config.MAX_TEXT_LENGTH);
  }

  return text;
}

/**
 * Extract text from a file by converting it to a temporary Google Doc via Drive API v3.
 * Works for PDFs (with OCR), DOCX, and images.
 * The temp doc is deleted after extraction.
 *
 * @param {File} file - The Drive file to extract text from
 * @param {Object} optionalArgs - Options passed to Drive.Files.create
 * @returns {string} Extracted text content
 */
function extractViaConversion(file, optionalArgs = {}) {
  const resource = {
    name: 'TEMP_EXTRACT_' + file.getName(),
    mimeType: 'application/vnd.google-apps.document',
  };

  // Drive API v3 handles creation and conversion implicitly via mimeType
  const tempFile = Drive.Files.create(resource, file.getBlob(), optionalArgs);

  try {
    const doc = DocumentApp.openById(tempFile.id);
    return doc.getBody().getText();
  } finally {
    // Standard DriveApp handles deletion safely across API versions
    if (tempFile && tempFile.id) {
      Drive.Files.remove(tempFile.id); // Permanently delete temp file instead of sending to Trash
    }
  }
}


// ── GEMINI CLASSIFICATION ──────────────────────────────────

/**
 * Send document text to Gemini for classification.
 * Returns a category key from getCategoryMap() or 'unknown'.
 */
function classifyWithGemini(text, fileName, config) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error('Gemini API key not set. Run setApiKey() first.');
  }

  const categories = Object.keys(getCategoryMap());
  const categoryList = categories.join(', ');

  const prompt = `You are a document classification system for home ownership documents. Classify the following document into exactly ONE of these categories:

${categoryList}

Rules:
- Respond with ONLY the category name, nothing else
- If you cannot confidently classify the document, respond with: unknown
- Base your decision on the document content, not just the filename
- Classification guide:
  - Purchase agreements, sales contracts, addenda, closing disclosures, HUD-1 statements, settlement statements, deeds, title insurance, title commitments, home inspection reports, property appraisals, loan estimates, promissory notes → purchase_title
  - Monthly mortgage statements, annual escrow analysis, refinance paperwork, loan modification documents → mortgage_finance
  - Homeowners/hazard insurance policies, flood or windstorm insurance, insurance claims, coverage declarations → insurance
  - Property tax bills, county tax statements, homestead exemption forms, tax protest filings, ARB decisions → tax_assessment
  - HOA declarations, CC&Rs, bylaws, rules, board meeting minutes, HOA dues invoices, special assessments, HOA correspondence → hoa
  - HVAC service records, plumbing invoices, electrical work, roof inspection/repairs, appliance repairs, general maintenance records, building permits, contractor bids/agreements/lien waivers, architectural plans/blueprints, home improvement receipts → maintenance_improvements
  - Product warranties, builder warranties, home warranty service contracts → warranty
  - Utility account setup, provider info, electric/gas/water/internet/trash bills → utilities_services

Document:
---
${text}
---

Category:`;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

  const payload = {
    contents: [{
      parts: [{
        text: prompt,
      }],
    }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 800,
    },
    safetySettings: [
      { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
    ]
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    Logger.log('   [ERROR] Gemini API error (' + responseCode + '): ' + response.getContentText());
    return 'unknown';
  }

  const result = JSON.parse(response.getContentText());

  if (!result.candidates || !result.candidates[0] || !result.candidates[0].content) {
    return 'unknown';
  }

  // Use modern optional chaining to safely extract text
  const raw = result?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!raw) {
    Logger.log('   [WARNING] Gemini response missing parts[0].text');
    Logger.log('   [DEBUG] Full API Response: ' + JSON.stringify(result, null, 2));
    return 'unknown';
  }

  const classification = raw.trim().toLowerCase();

  // Exact match only — no fuzzy matching to prevent misclassification
  const validCategories = Object.keys(getCategoryMap());
  if (validCategories.includes(classification)) {
    return classification;
  }

  Logger.log('   [WARNING] Gemini returned unrecognized category: "' + classification + '"');
  return 'unknown';
}


// ── FILE OPERATIONS ────────────────────────────────────────

/**
 * Copy a file to the destination folder with the -cp suffix.
 * Returns the new file reference.
 *
 * Examples:
 * "settlement-statement.pdf"  → "settlement-statement-cp.pdf"
 * "DOC001.pdf"                → "DOC001-cp.pdf"
 */
function copyFileToFolder(file, destFolder, suffix) {
  const originalName = file.getName();
  const newName = addSuffixToFilename(originalName, suffix);

  const copy = file.makeCopy(newName, destFolder);
  return copy;
}

/**
 * Move a file to the destination folder using the modern Drive API.
 */
function moveFileToFolder(file, destFolder) {
  file.moveTo(destFolder);
}

/**
 * Insert suffix before the file extension.
 */
function addSuffixToFilename(filename, suffix) {
  const lastDot = filename.lastIndexOf('.');
  if (lastDot === -1) {
    return filename + suffix;
  }
  const name = filename.substring(0, lastDot);
  const ext = filename.substring(lastDot);
  return name + suffix + ext;
}

/**
 * Get a sub-folder by name within a parent folder.
 */
function getSubFolder(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return null;
}


// ── MANUAL UTILITIES ───────────────────────────────────────

/**
 * Run this to manually process right now
 * instead of waiting for the trigger.
 */
function processNow() {
  processOriginals();
}

/**
 * Run this to test classification on a single file.
 * Change the FILE_ID to test with a specific file.
 */
function testClassification() {
  const FILE_ID = 'PASTE_A_FILE_ID_HERE';

  if (FILE_ID === 'PASTE_A_FILE_ID_HERE') {
    Logger.log('[ERROR] Replace FILE_ID with an actual file ID first');
    return;
  }

  const config = getConfig();

  if (!config.ROOT_FOLDER_ID) {
    Logger.log('[ERROR] ROOT_FOLDER_ID not set in getConfig(). Run setupFolderStructure() first.');
    return;
  }

  const file = DriveApp.getFileById(FILE_ID);
  const text = extractText(file, config);

  Logger.log('Extracted text (first 500 chars):');
  Logger.log(text.substring(0, 500));
  Logger.log('');

  const classification = classifyWithGemini(text, file.getName(), config);
  Logger.log('Classification: ' + classification);

  const categoryMap = getCategoryMap();
  Logger.log('Destination: ' + (categoryMap[classification] || 'UNSORTED'));
}

/**
 * Run this to see all folder IDs in your structure.
 */
function listFolderIds() {
  const config = getConfig();

  if (!config.ROOT_FOLDER_ID) {
    Logger.log('[ERROR] ROOT_FOLDER_ID not set in getConfig(). Run setupFolderStructure() first.');
    return;
  }

  const rootFolder = DriveApp.getFolderById(config.ROOT_FOLDER_ID);

  Logger.log('Root: ' + rootFolder.getName() + ' (' + rootFolder.getId() + ')');

  const folders = rootFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    Logger.log('  ' + folder.getName() + ' (' + folder.getId() + ')');
  }
}

/**
 * Run this to see processing history —
 * which files were processed and how they were classified.
 */
function viewProcessingHistory() {
  const processed = getProcessedFileIds();
  const entries = Object.entries(processed);

  Logger.log('[INFO] Processing history: ' + entries.length + ' files');
  Logger.log('');

  for (const [fileId, info] of entries) {
    try {
      const file = DriveApp.getFileById(fileId);
      Logger.log('  ' + file.getName() + ' → ' + info.classification + ' (' + info.processedAt + ')');
    } catch (e) {
      Logger.log('  [deleted file ' + fileId + '] → ' + info.classification + ' (' + info.processedAt + ')');
    }
  }
}