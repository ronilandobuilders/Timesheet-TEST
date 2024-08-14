var CLIENT_ID = '1000614174080-gkrie47bmbokh4v024rqqphimtde4uf0.apps.googleusercontent.com';
var CLIENT_SECRET = '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';

function getService() {
return createService('MyServiceAccount')
.setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
.setTokenUrl('https://accounts.google.com/o/oauth2/token')
.setClientId(CLIENT_ID)
.setClientSecret(CLIENT_SECRET)
.setCallbackFunction('authCallback')
.setPropertyStore(PropertiesService.getScriptProperties())
.setScope('https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/script.scriptapp https://www.googleapis.com/auth/drive');
}

function authCallback(request) {
var service = getService();
if (service.hasAccess()) {
var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + sheet.getId() + '/authCallback';
var response = UrlFetchApp.fetch(url, {
headers: {
'Authorization': 'Bearer ' + service.getAccessToken()
}
});
var data = JSON.parse(response.getContentText());
Logger.log(data);
} else {
Logger.log('Access not granted');
}
}

function setupLockTriggers() {
var service = getService();
if (service.hasAccess()) {
// Clear existing triggers to avoid duplicates
var triggers = ScriptApp.getProjectTriggers();
triggers.forEach(function(trigger) {
if (trigger.getHandlerFunction() == 'autoLockAC' || trigger.getHandlerFunction() == 'autoLockDE') {
ScriptApp.deleteTrigger(trigger);
}
});

// Create new triggers for the new sheet
var service = getService();
if (service.hasAccess()) {
var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + sheet.getId() + '/triggers';
var response = UrlFetchApp.fetch(url, {
headers: {
'Authorization': 'Bearer ' + service.getAccessToken()
}
});
var data = JSON.parse(response.getContentText());
Logger.log(data);
} else {
Logger.log('Access not granted');
}

} else {
Logger.log('Authorization error: ' + service.getLastError());
}
}

function copyTemplateWithPermissions() {
var service = getService();
if (service.hasAccess()) {
var templateId = '1GHawh-mJvT-n0jmfWlh5wlLpzP_wCmwtz_qsH8eoGNk'; // Your template sheet ID

// Calculate the start and end dates of the current week (Sunday to Saturday)
var now = new Date();
var dayOfWeek = now.getDay();
var startOfWeek = new Date(now);
startOfWeek.setDate(now.getDate() - dayOfWeek); // Sunday
var endOfWeek = new Date(startOfWeek);
endOfWeek.setDate(startOfWeek.getDate() + 6); // Saturday

// Format the dates to "MM/DD"
var startDateStr = (startOfWeek.getMonth() + 1) + '/' + startOfWeek.getDate();
var endDateStr = (endOfWeek.getMonth() + 1) + '/' + endOfWeek.getDate();

var newSheetName = 'Timesheet Report ' + startDateStr + ' - ' + endDateStr;

// Copy the template sheet
var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + templateId + '/copy';
var response = UrlFetchApp.fetch(url, {
headers: {
'Authorization': 'Bearer ' + service.getAccessToken()
}
});
var data = JSON.parse(response.getContentText());
var newSheetId = data.spreadsheetId;
var newSheet = SpreadsheetApp.openById(newSheetId);
Utilities.sleep(5000); // Wait for 5 seconds

var sheetNames = newSheet.getSheets().map(function(sheet) {
return sheet.getName();
});
Logger.log('New sheet created with sheets: ' + sheetNames.join(', '));

applyPermissions(newSheet);
addLockUnlockMenu(newSheet);
setupLockTriggers(); // Ensure this is called correctly

Logger.log('New sheet created with permissions set: ' + newSheet.getUrl());
}

function applyPermissions(newSheet) {
var permissions = [
{ sheetName: 'Danny', email: 'pulliam.daniel0508@gmail.com' },
{ sheetName: 'David M.', email: 'deeloksz@gmail.com' },
{ sheetName: 'David P.', email: 'davidpriest493@gmail.com' },
{ sheetName: 'Jake', email: 'jakehernandez2020@gmail.com' },
{ sheetName: 'Justin', email: 'wayne.allcoastpa@gmail.com' },
{ sheetName: 'Anthony', email: 'areynel88@gmail.com' },
{ sheetName: 'Arvon', email: 'arvongregory33@gmail.com' },
{ sheetName: 'Radik', email: 'radikshastitko1984@gmail.com' },
];

var projectManagers = ['greg@landobuilders.com', 'alexander@landobuilders.com'];
var ownerEmail = 'roni@landobuilders.com';
permissions.forEach(function(permission) {
Logger.log('Applying permissions for sheet: ' + permission.sheetName);
var sheet = newSheet.getSheetByName(permission.sheetName);
if (sheet) {
Logger.log('Sheet found: ' + permission.sheetName);
var service = getService();
if (service.hasAccess()) {
var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + sheet.getId() + '/protections';
var response = UrlFetchApp.fetch(url, {
headers: {
'Authorization': 'Bearer ' + service.getAccessToken()
}
});
var data = JSON.parse(response.getContentText());
Logger.log(data);
} else {
Logger.log('Access not granted');
}

} else {
Logger.log('Sheet not found: ' + permission.sheetName);
}
});
}

function getAuthorizationUrl() {
var service = getService();
if (!service.hasAccess()) {
var authorizationUrl = service.getAuthorizationUrl();
Logger.log('Open the following URL and complete the authorization flow: ' + authorizationUrl);
} else {
Logger.log('Access already granted');
}
}
}

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

function setupLockTriggers() {
var service = getService();
if (service.hasAccess()) {
// Clear existing triggers to avoid duplicates
var triggers = ScriptApp.getProjectTriggers();
triggers.forEach(function(trigger) {
if (trigger.getHandlerFunction() == 'autoLockAC' || trigger.getHandlerFunction() == 'autoLockDE') {
ScriptApp.deleteTrigger(trigger);
}
});

// Create new triggers for the new sheet
ScriptApp.newTrigger('autoLockAC')
.timeBased()
.onWeekDay(ScriptApp.WeekDay.THURSDAY)
.atHour(9)
.inTimezone('America/Los_Angeles')
.create();

ScriptApp.newTrigger('autoLockDE')
.timeBased()
.onWeekDay(ScriptApp.WeekDay.MONDAY)
.atHour(9)
.inTimezone('America/Los_Angeles')
.create();
} else {
Logger.log('Authorization error: ' + service.getLastError());
}
}

function autoLockAC() {
Logger.log('Starting autoLockAC');
lockAC();
Logger.log('Completed autoLockAC');
}

function autoLockDE() {
Logger.log('Starting autoLockDE');
lockDE();
Logger.log('Completed autoLockDE');
}

function testAutoLockAC() {
lockAC();
}

function testAutoLockDE() {
lockDE();
}

function testLockCarlos() {
Logger.log('Starting testLockCarlos');
var sheetName = 'Carlos';
var email = 'roni@landobuilders.com';

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
} catch (e) {
Logger.log('Error locking sheet ' + sheetName + ': ' + e.toString());
}
} else {
Logger.log('Sheet not found: ' + sheetName);
}
Logger.log('Completed testLockCarlos');
}

function testLockDanny() {
Logger.log('Starting testLockDanny');
var sheetName = 'Danny';
var email = 'roni@landobuilders.com';

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
} catch (e) {
Logger.log('Error locking sheet ' + sheetName + ': ' + e.toString());
}
} else {
Logger.log('Sheet not found: ' + sheetName);
}
Logger.log('Completed testLockDanny');
}

function testLockDavidM() {
Logger.log('Starting testLockDavidM');
var sheetName = 'David M.';
var email = 'roni@landobuilders.com';

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
} catch (e) {
Logger.log('Error locking sheet ' + sheetName + ': ' + e.toString());
}
} else {
Logger.log('Sheet not found: ' + sheetName);
}
Logger.log('Completed testLockDavidM');
}

function testLockDavidP() {
Logger.log('Starting testLockDavidP');
var sheetName = 'David P.';
var email = 'roni@landobuilders.com';

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
} catch (e) {
Logger.log('Error locking sheet ' + sheetName + ': ' + e.toString());
}
} else {
Logger.log('Sheet not found: ' + sheetName);
}
Logger.log('Completed testLockDavidP');
}

function testLockJake() {
Logger.log('Starting testLockJake');
var sheetName = 'Jake';
var email = 'roni@landobuilders.com';

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
} catch (e) {
Logger.log('Error locking sheet ' + sheetName + ': ' + e.toString());
}
} else {
Logger.log('Sheet not found: ' + sheetName);
}
Logger.log('Completed testLockJake');
}

function testLockJustin() {
Logger.log('Starting testLockJustin');
var sheetName = 'Justin';
var email = 'roni@landobuilders.com';

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
} catch (e) {
Logger.log('Error locking sheet ' + sheetName + ': ' + e.toString());
}
} else {
Logger.log('Sheet not found: ' + sheetName);
}
Logger.log('Completed testLockJustin');
}

