// File Request Sheet ID and Name
const FILE_REQUEST_SHEET_ID = '1vt2AjI1DPUpHGQuCnP_Lhqy-ZvPmoYi1JzmPpSiE8So';
const USERS_SHEET_ID = '1_2PKNZfCCVFYCZ8RUUznDFpRpw-po6dKIYse3a8rjzk';
const FILE_REQUEST_SHEET_NAME = 'Requests';
const USERS_SHEET_NAME = 'Users';

// Serve the HTML frontend
function doGet(e) {
  try {
    // Check if requesting main app page
    const page = e && e.parameter && e.parameter.page;
    
    if (page === 'main') {
      return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('File Request System')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // Default: Serve login page
    return HtmlService.createTemplateFromFile('login')
      .evaluate()
      .setTitle('File Request System - Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return HtmlService.createHtmlOutput("Error loading app: " + err);
  }
}


// Include function for HTML templates
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doPost(e) {
  try {
    // Accept FormData (multipart/form-data) and get fields from e.parameter or parse from e.postData
    let data = e.parameter || {};
    Logger.log('doPost received:');
    Logger.log(JSON.stringify({ parameter: e.parameter, postData: e.postData }));
    
    if (e.postData && e.postData.type && e.postData.type.indexOf('multipart') !== -1) {
      // Parse multipart form data
      const parts = e.postData.contents.split('&');
      data = {};
      for (let i = 0; i < parts.length; i++) {
        const kv = parts[i].split('=');
        if (kv.length === 2) {
          data[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1].replace(/\+/g, ' '));
        }
      }
      Logger.log('Parsed multipart data: ' + JSON.stringify(data));
    }
    
    const action = data.action || 'submitFileRequest'; // Default action for backward compatibility
    Logger.log('Action: ' + action);
    
    let result;
    if (action === 'submitFileRequest') {
      result = submitFileRequest(data);
      Logger.log('Result: ' + JSON.stringify(result));
      return createResponse(result);
    }
    
    if (action === 'getFileRequests') {
      result = getFileRequests();
      Logger.log('Result: ' + JSON.stringify(result));
      return createResponse(result);
    }
    
    if (action === 'login') {
      result = authenticateUser(data.username, data.password);
      Logger.log('Result: ' + JSON.stringify(result));
      return createResponse(result);
    }
    
    if (action === 'signup') {
      result = createUser(data);
      Logger.log('Result: ' + JSON.stringify(result));
      return createResponse(result);
    }
    
    if (action === 'updateRequestStatus') {
      result = updateRequestStatus(data);
      Logger.log('Result: ' + JSON.stringify(result));
      return createResponse(result);
    }
    Logger.log('Unknown action: ' + action);
    return createResponse({ success: false, error: 'Unknown action' });
// Update request status by ID
function updateRequestStatus(data) {
  try {
    const ss = SpreadsheetApp.openById(FILE_REQUEST_SHEET_ID);
    const sheet = ss.getSheetByName(FILE_REQUEST_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    const id = data.id;
    const status = data.status || '';
    const remarks = data.remarks || '';
    if (!id) return { success: false, error: 'No ID provided' };
    const dataRows = sheet.getDataRange().getValues();
    for (let i = 1; i < dataRows.length; i++) {
      if (String(dataRows[i][0]) === String(id)) {
        sheet.getRange(i + 1, 9).setValue(status); // STATUS OF REQUEST (c/o miss Jane)
        if (remarks) {
          sheet.getRange(i + 1, 10).setValue(remarks); // REMARKS
        }
        // Save Received by and DATE RECEIVED if provided
        if (data["Received by"]) {
          sheet.getRange(i + 1, 4).setValue(data["Received by"]); // 4th column: Received by
        }
        if (data["DATE RECEIVED"]) {
          sheet.getRange(i + 1, 5).setValue(data["DATE RECEIVED"]); // 5th column: DATE RECEIVED
        }
        // Save SIGNED COPY RECEIVED BY and RECEIVED DATE if provided
        if (data["SIGNED COPY RECEIVED BY"]) {
          sheet.getRange(i + 1, 11).setValue(data["SIGNED COPY RECEIVED BY"]); // 11th column
        }
        if (data["RECEIVED DATE"]) {
          sheet.getRange(i + 1, 12).setValue(data["RECEIVED DATE"]); // 12th column
        }
        return { success: true };
      }
    }
    return { success: false, error: 'ID not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
  } catch (err) {
    Logger.log('doPost error: ' + err + '\n' + (err.stack || ''));
    return createResponse({ success: false, error: String(err) });
  }
}

// File Request: Append a new row to FileRequests sheet
function submitFileRequest(data) {
  try {
    const ss = SpreadsheetApp.openById(FILE_REQUEST_SHEET_ID);
    let sheet = ss.getSheetByName(FILE_REQUEST_SHEET_NAME);
    if (!sheet) {
      // Create sheet with headers if not exists
      sheet = ss.insertSheet(FILE_REQUEST_SHEET_NAME);
      sheet.appendRow([
        'ID', 'Date Submitted', 'Forwarded by', 'Received by', 'DATE RECEIVED', 
        'Name of Company/Contractor', 'Amount (if Billing)', 'PARTICULAR/S (content of the request)', 
        'STATUS OF REQUEST (c/o miss Jane)', 'REMARKS'
      ]);
    }
    // Auto-increment ID
    const lastRow = sheet.getLastRow();
    let nextId = 1;
    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue();
      nextId = Number(lastId) + 1;
    }
    const now = new Date();
    const dateSubmitted = now.getFullYear() + '-' + 
                         String(now.getMonth() + 1).padStart(2, '0') + '-' + 
                         String(now.getDate()).padStart(2, '0') + ' ' +
                         String(now.getHours()).padStart(2, '0') + ':' + 
                         String(now.getMinutes()).padStart(2, '0') + ':' + 
                         String(now.getSeconds()).padStart(2, '0') + ' ' +
                         (now.getHours() >= 12 ? 'PM' : 'AM');
    
    // Auto-fill Received by with user's first name if available
    let receivedBy = '';
    if (data.userFirstName) {
      receivedBy = data.userFirstName;
    } else if (data.receivedBy) {
      receivedBy = data.receivedBy;
    }
    sheet.appendRow([
      nextId,                              // ID
      dateSubmitted,                       // Date Submitted
      data.forwardedBy || '',             // Forwarded by
      receivedBy,                         // Received by
      data.dateReceived || '',            // DATE RECEIVED
      data.companyName || '',             // Name of Company/Contractor
      data.amount || '',                  // Amount (if Billing) - keep as formatted string
      data.particulars || '',             // PARTICULAR/S (content of the request)
      data.status || 'pending',           // STATUS OF REQUEST (c/o miss Jane)
      data.remarks || ''                  // REMARKS
    ]);
    return { success: true, id: nextId, message: 'File request submitted successfully!' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// File Request: Fetch all file requests
function getFileRequests() {
  try {
    const ss = SpreadsheetApp.openById(FILE_REQUEST_SHEET_ID);
    const sheet = ss.getSheetByName(FILE_REQUEST_SHEET_NAME);
    if (!sheet) {
      return { success: true, fileRequests: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const fileRequests = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fileRequest = {};
      for (let j = 0; j < headers.length; j++) {
        fileRequest[headers[j]] = row[j];
      }
      fileRequests.push(fileRequest);
    }
    
    return { success: true, fileRequests: fileRequests };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// User Authentication (login)
function authenticateUser(username, password) {
  try {
    const ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!sheet) {
      sheet = createUsersSheet(ss);
    }
    const data = sheet.getDataRange().getValues();
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // username: col 4, password: col 5, status: col 8, office: col 3
      if (row[4] && row[5] && row[8]) {
        const allowedOffices = ['pao', 'cfpdo', 'admin'];
        const userOffice = (row[3] || '').toString().toLowerCase().trim();
        const userStatus = (row[8] || '').toString().toLowerCase().trim();
        
        Logger.log('Checking user:', {
          username: row[4].toString(),
          office: userOffice,
          status: userStatus,
          inputUsername: username,
          inputPassword: password ? '***' : 'empty'
        });
        
        if (
          row[4].toString().trim() === username.trim() &&
          row[5].toString().trim() === password.trim() &&
          userStatus === 'active' &&
          allowedOffices.includes(userOffice)
        ) {
          return {
            success: true,
            user: {
              userId: row[0],
              firstName: row[1],
              lastName: row[2],
              fullName: (row[1] || '') + ' ' + (row[2] || ''),
              office: row[3],
              username: row[4],
              userType: row[6],
              email: row[6],
              phone: row[7],
              status: row[8]
            }
          };
        } else {
          // Check why login failed for debugging
          if (row[4].toString().trim() === username.trim()) {
            if (row[5].toString().trim() !== password.trim()) {
              Logger.log('Password mismatch for user:', username);
            } else if (userStatus !== 'active') {
              Logger.log('User not active:', username, 'Status:', userStatus);
              return { success: false, error: 'Account is not active. Please contact administrator.' };
            } else if (!allowedOffices.includes(userOffice)) {
              Logger.log('Office not allowed:', userOffice);
              return { success: false, error: 'Your office is not authorized to access this system.' };
            }
          }
        }
      }
    }
    
    Logger.log('No matching user found for:', username);
    return { success: false, error: 'Invalid username or password' };
  } catch (error) {
    Logger.log('Authentication error:', error);
    return { success: false, error: error.toString() };
  }
}

// Create new user (signup)
function createUser(data) {
  try {
    const ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!sheet) {
      sheet = createUsersSheet(ss);
    }
    
    // Check if username already exists
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      const row = existingData[i];
      if (row[4] && row[4].toString() === data.username) {
        return { success: false, error: 'Username already exists' };
      }
    }
    
    // Generate new user ID
    const lastRow = sheet.getLastRow();
    let nextId = 1;
    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue();
      nextId = Number(lastId) + 1;
    }
    
    // Add new user
    sheet.appendRow([
      nextId,                           // User ID
      data.firstName || '',             // First Name
      data.lastName || '',              // Last Name
      data.office || '',                // Office
      data.username || '',              // username
      data.password || '',              // password
      data.email || '',                 // email
      data.phone || '',                 // phone
      'Active'                          // Status
    ]);
    
    return { 
      success: true, 
      message: 'User created successfully!',
      userId: nextId 
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function createUsersSheet(ss) {
  const sheet = ss.insertSheet(USERS_SHEET_NAME);
  // Add headers: User ID, First Name, Last Name, Office, username, password, email, phone, Status
  sheet.appendRow(['User ID', 'First Name', 'Last Name', 'Office', 'username', 'password', 'email', 'phone', 'Status']);
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, 9);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f4f6');
  headerRange.setBorder(true, true, true, true, true, true);
  // Set column widths
  sheet.setColumnWidth(1, 200); // User ID
  sheet.setColumnWidth(2, 120); // First Name
  sheet.setColumnWidth(3, 120); // Last Name
  sheet.setColumnWidth(4, 150); // Office
  sheet.setColumnWidth(5, 120); // username
  sheet.setColumnWidth(6, 120); // password
  sheet.setColumnWidth(7, 200); // email
  sheet.setColumnWidth(8, 120); // phone
  sheet.setColumnWidth(9, 100); // Status
  // Add default admin user (User ID 1)
  sheet.appendRow([
    1,
    'System',
    'Administrator',
    'PAO BASAK',
    'admin',
    'admin',
    'admin@reservehub.com',
    '+1 (555) 123-4567',
    'Active'
  ]);
  return sheet;
}

// Preflight for CORS (not needed for same-origin, but keeping for compatibility)
function doOptions(e) {
  return createResponse({});
}

// Helper for JSON response (no CORS needed - same domain!)
function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}