function addLockUnlockMenu(sheet) {
var ui = SpreadsheetApp.getUi();
ui.createMenu('Lock/Unlock')
.addItem('Lock A - C', 'lockAC')
.addItem('Unlock A - C', 'unlockAC')
.addItem('Lock D - E', 'lockDE')
.addItem('Unlock D - E', 'unlockDE')
.addToUi();
}

function onOpen() {
addLockUnlockMenu(SpreadsheetApp.getActiveSpreadsheet());
}

function lockAC() {
Logger.log('Starting lockAC');
var sheets = ['Danny', 'David M.', 'David P.', 'Jake', 'Justin', 'Anthony', 'Arvon', 'Radik'];
var email = 'roni@landobuilders.com';
sheets.forEach(function(sheetName) {
Logger.log('Locking sheet: ' + sheetName);
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
if (sheet) {
try {
Logger.log('Getting range A:C for sheet: ' + sheetName);
var range = sheet.getRange('A:C');
Logger.log('Protecting range A:C for sheet: ' + sheetName);
var protection = range.protect().setDescription('Protected Range A:C');
Logger.log('Removing editors from range A:C for sheet: ' + sheetName);
protection.removeEditors(protection.getEditors());
Logger.log('Adding editor to range A:C for sheet: ' + sheetName);
protection.addEditor(email);
Logger.log('Locked A:C for sheet: ' + sheetName);
// Commenting out the notification for debugging
// notifyLocked('A - C');
} catch (e) {
Logger.log('Error locking sheet ' + sheetName + ': ' + e.toString());
}
} else {
Logger.log('Sheet not found: ' + sheetName);
}
// Adding a delay to see if it's a timing issue
Utilities.sleep(2000); // Sleep for 2 seconds
});
Logger.log('Completed lockAC');
}

function unlockAC() {
var sheets = ['Danny', 'David M.', 'David P.', 'Jake', 'Justin', 'Anthony', 'Arvon', 'Radik'];
sheets.forEach(function(sheetName) {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
if (sheet) {
var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
protections.forEach(function(protection) {
var range = protection.getRange();
var rangeNotation = range.getA1Notation();
if (rangeNotation.startsWith('A') || rangeNotation.startsWith('B') || rangeNotation.startsWith('C')) {
if (rangeNotation !== 'A1:E1' && rangeNotation !== 'A2:E30' && rangeNotation !== 'A31:E50') {
protection.remove();
}
}
});
}
});
}

function lockDE() {
Logger.log('Starting lockDE');
var sheets = ['Danny', 'David M.', 'David P.', 'Jake', 'Justin', 'Anthony', 'Arvon', 'Radik'];
var email = 'roni@landobuilders.com';
sheets.forEach(function(sheetName) {
Logger.log('Locking sheet: ' + sheetName);
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
if (sheet) {
try {
Logger.log('Getting range D:E for sheet: ' + sheetName);
var range = sheet.getRange('D:E');
Logger.log('Protecting range D:E for sheet: ' + sheetName);
var protection = range.protect().setDescription('Protected Range D:E');
Logger.log('Removing editors from range D:E for sheet: ' + sheetName);
protection.removeEditors(protection.getEditors());
Logger.log('Adding editor to range D:E for sheet: ' + sheetName);
protection.addEditor(email);
Logger.log('Locked D:E for sheet: ' + sheetName);
} catch (e) {
Logger.log('Error locking sheet ' + sheetName + ': ' + e.toString());
}
} else {
Logger.log('Sheet not found: ' + sheetName);
}
// Adding a delay to see if it's a timing issue
Utilities.sleep(2000); // Sleep for 2 seconds
});
Logger.log('Completed lockDE');
}

function notifyLocked(sections) {
var ui = SpreadsheetApp.getUi();
ui.alert('Sections ' + sections + ' have been locked at 9:00 AM!');
}

function unlockDE() {
var sheets = ['Danny', 'David M.', 'David P.', 'Jake', 'Justin', 'Anthony', 'Arvon', 'Radik'];
sheets.forEach(function(sheetName) {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
if (sheet) {
var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
protections.forEach(function(protection) {
var range = protection.getRange();
var rangeNotation = range.getA1Notation();
if (rangeNotation.startsWith('D') || rangeNotation.startsWith('E')) {
if (rangeNotation !== 'A1:E1' && rangeNotation !== 'A2:E30' && rangeNotation !== 'A31:E50') {
protection.remove();
}
}
});
}
});
}

function notifyLocked(sections) {
var ui = SpreadsheetApp.getUi();
ui.alert('Sections ' + sections + ' have been locked at 9:00 AM!');
}
