/**
 * Creates two time-driven triggers.
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
function createTimeDrivenTriggers() {

    // Trigger every {Enter Weekday} at {Enter Hour of Day}.
    ScriptApp.newTrigger('sendDodahsDateCheckEmail')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.THURSDAY)
        .atHour(4)
        .create();
    
}
