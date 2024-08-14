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
