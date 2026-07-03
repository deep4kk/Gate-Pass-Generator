/**
 * Gate Pass Generator
 * Google Apps Script Web Application
 */

const CONFIG = {
  SHEET_NAMES: {
    USERS: 'Users',
    GATE_PASSES: 'GatePasses',
    AUDIT_LOG: 'AuditLog'
  },
  PASS_TYPES: ['Visitor', 'Material', 'Vehicle', 'Emergency'],
  STATUS: { PENDING: 'Pending', APPROVED: 'Approved', REJECTED: 'Rejected', USED: 'Used', EXPIRED: 'Expired' }
};

const SHEET_CONFIG = const SHEET_CONFIG = {
  Users: { columns: ['Email', 'Name', 'Role', 'Department', 'Active'], sampleData: [['admin@company.com', 'Admin User', 'Admin', 'Security', true]] },
  GatePasses: { columns: ['PassID', 'Type', 'RequesterEmail', 'RequesterName', 'VisitorName', 'Purpose', 'Destination', 'VehicleNumber', 'ValidFrom', 'ValidUntil', 'Status', 'QRData', 'CreatedAt', 'ApprovedBy', 'ApprovedAt'] },
  AuditLog: { columns: ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue'] }
};;

// ============== ENHANCED UTILITIES (v2.0) ==============
const VERSION = '2.0.0';

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  for (const [sheetName, config] of Object.entries(SHEET_CONFIG)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      created.push(sheetName);
      sheet.getRange(1, 1, 1, config.columns.length).setValues([config.columns]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      if (config.sampleData) config.sampleData.forEach(row => sheet.appendRow(row));
    }
  }
  Logger.log('InitializeSheets: Created ' + created.join(', '));
  return { created };
}

function handleError(error, context) {
  const errorMsg = error instanceof Error ? error.message : String(error);
  Logger.log('[ERROR] ' + context + ': ' + errorMsg);
  if (typeof logAction === 'function') logAction('ERROR', context, '', errorMsg);
  return { success: false, error: errorMsg };
}

function backupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupName = 'Backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  ss.copy(backupName);
  if (typeof logAction === 'function') logAction('BACKUP_CREATED', 'System', '', backupName);
  return { success: true, backupName };
}

function exportToPDF(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  const pdf = DriveApp.getFileById(ss.getId()).getAs('application/pdf');
  const pdfName = sheet.getName() + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.pdf';
  DriveApp.getRootFolder().createFile(pdf).setName(pdfName);
  return { success: true, fileName: pdfName };
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 System Menu')
    .addItem('📊 Initialize Sheets', 'initializeSheets')
    .addItem('💾 Create Backup', 'backupData')
    .addItem('📄 Export to PDF', 'exportToPDF')
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('System v' + VERSION, 'Enhanced with:\n- initializeSheets()\n- backupData()\n- exportToPDF()', ui.ButtonSet.OK);
}

function getUserRole(email) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase() && data[i][4]) return data[i][2];
  }
  return null;
}

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase()) {
      return { email: data[i][0], name: data[i][1], role: data[i][2], department: data[i][3] };
    }
  }
  return { email, name: 'Unknown', role: null };
}

function requireRole(allowedRoles) {
  const user = getCurrentUser();
  if (!user.role || !allowedRoles.includes(user.role)) throw new Error('Unauthorized');
  return user;
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); setupHeaders(sheet, name); }
  return sheet;
}

function setupHeaders(sheet, name) {
  const headers = {
    'Users': ['Email', 'Name', 'Role', 'Department', 'Active'],
    'GatePasses': ['PassID', 'Type', 'RequesterEmail', 'RequesterName', 'VisitorName', 'Purpose', 
                   'Destination', 'VehicleNumber', 'ValidFrom', 'ValidUntil', 'Status', 'QRData', 
                   'CreatedAt', 'ApprovedBy', 'ApprovedAt'],
    'AuditLog': ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue']
  };
  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
  }
}

function generateId() { return 'GP-' + new Date().getFullYear() + '-' + Utilities.getUuid().substring(0, 6).toUpperCase(); }

function logAction(action, recordId, oldValue, newValue) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.AUDIT_LOG);
  sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action, recordId, oldValue, newValue]);
}

function sendNotification(toEmail, subject, htmlBody) {
  try { MailApp.sendEmail({ to: toEmail, subject, htmlBody }); return true; } catch (e) { return false; }
}

function generateQRCode(data) {
  const url = 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=' + encodeURIComponent(data);
  return url;
}

function createGatePass(passType, visitorName, purpose, destination, vehicleNumber, validFrom, validUntil) {
  const user = requireRole(['Employee', 'Manager', 'Admin']);
  
  const passId = generateId();
  const qrData = JSON.stringify({ passId, type: passType, valid: true });
  const qrUrl = generateQRCode(qrData);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.GATE_PASSES);
  sheet.appendRow([
    passId, passType, user.email, user.name, visitorName, purpose, destination, vehicleNumber || '',
    new Date(validFrom), new Date(validUntil), CONFIG.STATUS.PENDING, qrData, new Date(), '', ''
  ]);
  
  logAction('CREATE_GATE_PASS', passId, '', CONFIG.STATUS.PENDING);
  
  return { success: true, passId, qrUrl };
}

function approveGatePass(passId) {
  const user = requireRole(['Security', 'Admin']);
  const sheet = getSheet(CONFIG.SHEET_NAMES.GATE_PASSES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === passId) {
      const row = i + 1;
      sheet.getRange(row, 11, 1, 1).setValue(CONFIG.STATUS.APPROVED);
      sheet.getRange(row, 13, 1, 1).setValue(user.email);
      sheet.getRange(row, 14, 1, 1).setValue(new Date());
      logAction('APPROVE_PASS', passId, CONFIG.STATUS.PENDING, CONFIG.STATUS.APPROVED);
      return { success: true };
    }
  }
  throw new Error('Pass not found');
}

function rejectGatePass(passId, reason) {
  const user = requireRole(['Security', 'Admin']);
  const sheet = getSheet(CONFIG.SHEET_NAMES.GATE_PASSES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === passId) {
      const row = i + 1;
      sheet.getRange(row, 11, 1, 1).setValue(CONFIG.STATUS.REJECTED);
      logAction('REJECT_PASS', passId, CONFIG.STATUS.PENDING, CONFIG.STATUS.REJECTED);
      return { success: true };
    }
  }
  throw new Error('Pass not found');
}

function verifyPass(passId) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.GATE_PASSES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === passId) {
      const validFrom = new Date(data[i][8]);
      const validUntil = new Date(data[i][9]);
      const now = new Date();
      const isValid = data[i][10] === CONFIG.STATUS.APPROVED && now >= validFrom && now <= validUntil;
      
      return {
        valid: isValid,
        status: data[i][10],
        type: data[i][1],
        visitorName: data[i][4],
        purpose: data[i][5],
        validFrom: data[i][8],
        validUntil: data[i][9]
      };
    }
  }
  return { valid: false, error: 'Pass not found' };
}

function getMyPasses() {
  const user = getCurrentUser();
  const sheet = getSheet(CONFIG.SHEET_NAMES.GATE_PASSES);
  const data = sheet.getDataRange().getValues();
  const passes = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toLowerCase() === user.email.toLowerCase()) {
      passes.push({
        passId: data[i][0],
        type: data[i][1],
        visitorName: data[i][4],
        purpose: data[i][5],
        destination: data[i][6],
        vehicleNumber: data[i][7],
        validFrom: data[i][8],
        validUntil: data[i][9],
        status: data[i][10],
        qrCode: generateQRCode(data[i][11]),
        createdAt: data[i][12]
      });
    }
  }
  return passes;
}

function getPendingApprovals() {
  requireRole(['Security', 'Admin']);
  const sheet = getSheet(CONFIG.SHEET_NAMES.GATE_PASSES);
  const data = sheet.getDataRange().getValues();
  const passes = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][10] === CONFIG.STATUS.PENDING) {
      passes.push({
        passId: data[i][0],
        type: data[i][1],
        requesterName: data[i][3],
        requesterEmail: data[i][2],
        visitorName: data[i][4],
        purpose: data[i][5],
        destination: data[i][6],
        vehicleNumber: data[i][7],
        validFrom: data[i][8],
        validUntil: data[i][9],
        createdAt: data[i][12]
      });
    }
  }
  return passes;
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gate Pass Generator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
