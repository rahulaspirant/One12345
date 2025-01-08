function fetchCalendarEventsCustom(startDate, daysRange) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); 
  sheet.appendRow(['Calendar Name', 'Event Title', 'Start Time', 'End Time']); 

  const calendars = CalendarApp.getAllCalendars();
  if (!calendars || calendars.length === 0) {
    console.log('No calendars found.');
    return;
  }
  
const start = new Date("2024-10-02"); // Start date
const end = new Date(start.getTime() + 90 * 24 * 60 * 60 * 1000); // Custom range

  // Loop through each calendar and fetch events
  for (const calendar of calendars) {
    const events = calendar.getEvents(start, end);
    if (events.length > 0) {
      for (const event of events) {
        const startTime = event.getStartTime();
        const endTime = event.getEndTime();
        sheet.appendRow([calendar.getName(), event.getTitle(), startTime, endTime]);
      }
    }
  }
}
