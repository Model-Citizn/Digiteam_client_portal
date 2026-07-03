// ============================================================
// Model Citizn — Client Upload Portal Backend
// Deploy as: Web App → Execute as Me → Anyone can access
// ============================================================

const PARENT_FOLDER_ID = '1NBtYxmNMkx-DoSH8VL0jrdiiDlet4skY';
const TEAM_EMAIL = 'lance@modelcitizn.com';
const SPREADSHEET_ID = '1Mi5NEXs2FDQxInBJN_M9bilbAGsetA4ehwyYWavvbdc';
const LOG_TAB_NAME = 'Digital Team onboarding';

// ============================================================
// TEST FUNCTION — Run this first to verify Drive + Sheets work
// In the editor: select "testConnection" from the dropdown → Run
// Check the Execution Log for results
// ============================================================
function testConnection() {
  try {
    // Test 1: Can we access the parent folder?
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    Logger.log('✓ Parent folder found: ' + parentFolder.getName());

    // Test 2: Can we create a subfolder?
    const testFolderName = '_test_connection_' + Date.now();
    const testFolder = parentFolder.createFolder(testFolderName);
    Logger.log('✓ Test folder created: ' + testFolder.getUrl());

    // Test 3: Can we create a file in it?
    const testBlob = Utilities.newBlob('Hello from test', 'text/plain', 'test.txt');
    const testFile = testFolder.createFile(testBlob);
    Logger.log('✓ Test file created: ' + testFile.getUrl());

    // Test 4: Can we access the spreadsheet?
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(LOG_TAB_NAME);
    Logger.log('✓ Spreadsheet found: ' + ss.getName());
    Logger.log('✓ Tab "' + LOG_TAB_NAME + '" exists: ' + (sheet !== null));

    // Clean up test folder
    testFolder.setTrashed(true);
    Logger.log('✓ Test folder cleaned up');

    Logger.log('');
    Logger.log('=== ALL TESTS PASSED ===');
    Logger.log('Drive and Sheets connections are working.');
  } catch (err) {
    Logger.log('✗ ERROR: ' + err.toString());
    Logger.log('Fix this before deploying.');
  }
}

// ============================================================
// POST handler — receives file uploads from the portal
// ============================================================
function doPost(e) {
  try {
    Logger.log('doPost called');
    Logger.log('Has postData: ' + !!e.postData);
    Logger.log('Has parameter: ' + !!e.parameter);
    Logger.log('Parameter keys: ' + (e.parameter ? Object.keys(e.parameter).join(', ') : 'none'));

    // Read from form parameter (survives 302 redirect) or fall back to postData
    let raw;
    if (e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
      Logger.log('Reading from e.parameter.payload');
    } else if (e.postData && e.postData.contents) {
      raw = e.postData.contents;
      Logger.log('Reading from e.postData.contents');
    } else {
      throw new Error('No payload received');
    }

    const data = JSON.parse(raw);

    const clientName = data.clientName || 'Unknown';
    const clientEmail = data.clientEmail || '';
    const clientCompany = data.clientCompany || '';
    const categories = data.categories || '';
    const notes = data.notes || '';
    const fileName = data.fileName || 'unnamed_file';
    const fileType = data.fileType || 'application/octet-stream';
    const fileData = data.fileData || '';

    Logger.log('Upload from: ' + clientName + ' (' + clientEmail + ')');
    Logger.log('File: ' + fileName + ' (' + fileType + ')');
    Logger.log('Categories: ' + categories);

    // Get or create client subfolder using EMAIL as folder name
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const folderName = clientEmail || clientName;
    const existingFolders = parentFolder.getFoldersByName(folderName);

    let clientFolder;
    if (existingFolders.hasNext()) {
      clientFolder = existingFolders.next();
      Logger.log('Using existing folder: ' + folderName);
    } else {
      clientFolder = parentFolder.createFolder(folderName);
      Logger.log('Created new folder: ' + folderName);
    }

    // Decode and save file with date prefix
    const decoded = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(decoded, fileType, fileName);
    const datePrefix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    blob.setName(datePrefix + '_' + fileName);
    const savedFile = clientFolder.createFile(blob);

    Logger.log('File saved: ' + savedFile.getUrl());

    // Log to spreadsheet
    logUpload(clientName, clientEmail, clientCompany, categories, notes, fileName, savedFile.getUrl());

    // Email notification
    sendNotification(clientName, clientEmail, clientCompany, categories, notes, fileName, savedFile.getUrl());

    Logger.log('Upload complete');

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', fileUrl: savedFile.getUrl() })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Upload error: ' + err.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// GET handler — health check
// ============================================================
function doGet(e) {
  Logger.log('doGet called');

  // History lookup: ?action=history&email=client@example.com
  if (e && e.parameter && e.parameter.action === 'history') {
    const email = String(e.parameter.email || '').trim().toLowerCase();
    const out = { status: 'ok', uploads: [] };
    try {
      if (email) {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(LOG_TAB_NAME);
        if (sheet && sheet.getLastRow() > 1) {
          const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
          rows.forEach(function (r) {
            if (String(r[2]).trim().toLowerCase() === email) {
              out.uploads.push({
                timestamp: r[0] ? new Date(r[0]).toISOString() : '',
                clientName: String(r[1] || ''),
                categories: String(r[4] || ''),
                notes: String(r[5] || ''),
                fileName: String(r[6] || ''),
                fileUrl: String(r[7] || '')
              });
            }
          });
        }
      }
    } catch (err) {
      out.status = 'error';
      out.message = err.toString();
    }
    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'MC Upload Portal backend is running.' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Helpers
// ============================================================
function logUpload(name, email, company, categories, notes, fileName, fileUrl) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LOG_TAB_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_TAB_NAME);
    sheet.appendRow([
      'Timestamp', 'Client Name', 'Email', 'Company',
      'Categories', 'Notes', 'File Name', 'File URL'
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Client Name', 'Email', 'Company',
      'Categories', 'Notes', 'File Name', 'File URL'
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  sheet.appendRow([
    new Date(),
    name,
    email,
    company,
    categories,
    notes,
    fileName,
    fileUrl
  ]);

  Logger.log('Logged to spreadsheet');
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
  Logger.log('Notification sent to ' + TEAM_EMAIL);
}
