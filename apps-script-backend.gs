// ============================================================
// Model Citizn — Client Upload Portal Backend (hardened)
// Deploy as: Web App → Execute as Me → Anyone can access
// ============================================================

const PARENT_FOLDER_ID = '1NBtYxmNMkx-DoSH8VL0jrdiiDlet4skY';
const TEAM_EMAIL = 'lance@modelcitizn.com';
const SPREADSHEET_ID = '1Mi5NEXs2FDQxInBJN_M9bilbAGsetA4ehwyYWavvbdc';
const LOG_TAB_NAME = 'Digital Team onboarding';

// --- Security limits ---
const MAX_FILE_BYTES = 10 * 1024 * 1024;           // 10 MB per file (matches portal)
const MAX_UPLOADS_PER_WINDOW = 20;                 // per email
const RATE_WINDOW_SECONDS = 600;                   // 10 minutes
const ALLOWED_EXTENSIONS = ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'csv', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'pptx', 'txt', 'md', 'zip'];
const MAX_NAME_LEN = 100;
const MAX_NOTES_LEN = 2000;
const MAX_FILENAME_LEN = 180;

// ============================================================
// TEST FUNCTION — Run this first to verify Drive + Sheets work
// ============================================================
function testConnection() {
  try {
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    Logger.log('✓ Parent folder found: ' + parentFolder.getName());

    const testFolderName = '_test_connection_' + Date.now();
    const testFolder = parentFolder.createFolder(testFolderName);
    Logger.log('✓ Test folder created: ' + testFolder.getUrl());

    const testBlob = Utilities.newBlob('Hello from test', 'text/plain', 'test.txt');
    const testFile = testFolder.createFile(testBlob);
    Logger.log('✓ Test file created: ' + testFile.getUrl());

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(LOG_TAB_NAME);
    Logger.log('✓ Spreadsheet found: ' + ss.getName());
    Logger.log('✓ Tab "' + LOG_TAB_NAME + '" exists: ' + (sheet !== null));

    testFolder.setTrashed(true);
    Logger.log('✓ Test folder cleaned up');
    Logger.log('=== ALL TESTS PASSED ===');
  } catch (err) {
    Logger.log('✗ ERROR: ' + err.toString());
  }
}

// ============================================================
// Sanitizers / validators
// ============================================================
function isValidEmail(email) {
  return /^[^\s@]{1,64}@[^\s@]{1,255}\.[^\s@]{2,}$/.test(email);
}

function cleanText(value, maxLen) {
  // Strip control chars and angle brackets, collapse whitespace, cap length
  return String(value || '')
    .replace(/[\u0000-\u001f\u007f<>]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, maxLen);
}

function sanitizeFileName(name) {
  // Keep only the base name (no path separators), strip dangerous chars
  let n = String(name || 'unnamed_file');
  n = n.split(/[\/\\]/).pop();
  n = n.replace(/[\u0000-\u001f\u007f<>:"|?*]/g, '_').replace(/^\.+/, '_').trim();
  if (!n) n = 'unnamed_file';
  return n.slice(0, MAX_FILENAME_LEN);
}

function getExtension(fileName) {
  const parts = String(fileName).toLowerCase().split('.');
  return parts.length > 1 ? parts.pop() : '';
}

function checkRateLimit(email) {
  // Returns true if the sender is within limits
  const cache = CacheService.getScriptCache();
  const key = 'rl_' + email;
  const count = Number(cache.get(key) || 0);
  if (count >= MAX_UPLOADS_PER_WINDOW) return false;
  cache.put(key, String(count + 1), RATE_WINDOW_SECONDS);
  return true;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// POST handler — receives file uploads from the portal
// ============================================================
function doPost(e) {
  try {
    let raw;
    if (e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else if (e.postData && e.postData.contents) {
      raw = e.postData.contents;
    } else {
      throw new Error('No payload received');
    }

    const data = JSON.parse(raw);

    // --- Validate + sanitize all inputs ---
    const clientEmail = String(data.clientEmail || '').trim().toLowerCase();
    if (!isValidEmail(clientEmail)) {
      return jsonResponse({ status: 'error', message: 'A valid email address is required.' });
    }

    if (!checkRateLimit(clientEmail)) {
      return jsonResponse({ status: 'error', message: 'Too many uploads in a short time. Please wait a few minutes and try again.' });
    }

    const clientName = cleanText(data.clientName, MAX_NAME_LEN) || 'Unknown';
    const clientCompany = cleanText(data.clientCompany, MAX_NAME_LEN);
    const categories = cleanText(data.categories, 300);
    const notes = cleanText(data.notes, MAX_NOTES_LEN);
    const fileName = sanitizeFileName(data.fileName);
    const fileData = String(data.fileData || '');

    // --- File type allowlist ---
    const ext = getExtension(fileName);
    if (ALLOWED_EXTENSIONS.indexOf(ext) === -1) {
      return jsonResponse({ status: 'error', message: 'File type ".' + ext + '" is not allowed.' });
    }

    // --- Size cap (base64 is ~4/3 of raw size) ---
    if (fileData.length > MAX_FILE_BYTES * 1.37) {
      return jsonResponse({ status: 'error', message: 'File exceeds the 10 MB limit.' });
    }
    if (!fileData) {
      return jsonResponse({ status: 'error', message: 'No file data received.' });
    }

    // Never trust the client's MIME type — derive a safe generic one
    const fileType = 'application/octet-stream';

    // Get or create client subfolder using EMAIL as folder name
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const existingFolders = parentFolder.getFoldersByName(clientEmail);
    const clientFolder = existingFolders.hasNext()
      ? existingFolders.next()
      : parentFolder.createFolder(clientEmail);

    // Decode and save file with date prefix
    const decoded = Utilities.base64Decode(fileData);
    if (decoded.length > MAX_FILE_BYTES) {
      return jsonResponse({ status: 'error', message: 'File exceeds the 10 MB limit.' });
    }
    const blob = Utilities.newBlob(decoded, fileType, fileName);
    const datePrefix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    blob.setName(datePrefix + '_' + fileName);
    const savedFile = clientFolder.createFile(blob);

    logUpload(clientName, clientEmail, clientCompany, categories, notes, fileName, savedFile.getUrl());
    sendNotification(clientName, clientEmail, clientCompany, categories, notes, fileName, savedFile.getUrl());

    // Do not echo the Drive URL back to the browser
    return jsonResponse({ status: 'success' });

  } catch (err) {
    Logger.log('Upload error: ' + err.toString());
    // Generic message — don't leak internals to the client
    return jsonResponse({ status: 'error', message: 'Upload failed. Please try again.' });
  }
}

// ============================================================
// GET handler — health check + history lookup
// ============================================================
function doGet(e) {
  // History lookup: ?action=history&email=client@example.com
  if (e && e.parameter && e.parameter.action === 'history') {
    const email = String(e.parameter.email || '').trim().toLowerCase();
    const out = { status: 'ok', uploads: [] };
    try {
      if (email && isValidEmail(email)) {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(LOG_TAB_NAME);
        if (sheet && sheet.getLastRow() > 1) {
          const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
          rows.forEach(function (r) {
            if (String(r[2]).trim().toLowerCase() === email) {
              // NOTE: no fileUrl here — Drive links are never exposed publicly
              out.uploads.push({
                timestamp: r[0] ? new Date(r[0]).toISOString() : '',
                clientName: String(r[1] || ''),
                categories: String(r[4] || ''),
                notes: String(r[5] || ''),
                fileName: String(r[6] || '')
              });
            }
          });
        }
      }
    } catch (err) {
      Logger.log('History error: ' + err.toString());
      out.status = 'error';
      out.message = 'Could not load history.';
    }
    return jsonResponse(out);
  }

  return jsonResponse({ status: 'ok', message: 'MC Upload Portal backend is running.' });
}

// ============================================================
// Helpers
// ============================================================
function logUpload(name, email, company, categories, notes, fileName, fileUrl) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LOG_TAB_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_TAB_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Client Name', 'Email', 'Company',
      'Categories', 'Notes', 'File Name', 'File URL'
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  sheet.appendRow([new Date(), name, email, company, categories, notes, fileName, fileUrl]);
}

function sendNotification(name, email, company, categories, notes, fileName, fileUrl) {
  var subject = 'New upload from ' + name + (company ? ' (' + company + ')' : '');
  var body = [
    'New file uploaded via Client Portal',
    '',
    'Client: ' + name,
    'Email: ' + email,
    'Company: ' + (company || '—'),
    'Categories: ' + (categories || '—'),
    'Notes: ' + (notes || '—'),
    '',
    'File: ' + fileName,
    'Link: ' + fileUrl,
    '',
    '— Model Citizn Upload Portal'
  ].join('\n');

  MailApp.sendEmail(TEAM_EMAIL, subject, body);
}
