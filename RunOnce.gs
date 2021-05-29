function createBiWeeklyTrigger() {
  // Trigger every other TUESDAY at 9:00.
  ScriptApp.newTrigger('main').timeBased()
      .everyWeeks(2).onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .atHour(9)
      .create();
}
