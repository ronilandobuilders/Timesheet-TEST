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
