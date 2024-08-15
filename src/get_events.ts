function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Мои скрипты')
        .addItem('Запустить скрипт', 'getAllCalendarEvents')
        .addToUi();
}

function getAllCalendarEvents() {
    var calendars = CalendarApp.getAllCalendars(); // Получаем все календари аккаунта
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    var datenow = new Date(); // Начало интервала (сейчас)
    var startTimeBas = new Date();
    startTimeBas.setDate(datenow.getDate() - 1);
    var endTime = new Date();
    endTime.setDate(datenow.getDate() + 7); // Конец интервала (через 7 дней)
    
    sheet.clear(); // Очистка листа
    sheet.appendRow(["Calendar Name", "Event Title", "Start Time", "End Time", "Description", "Color", "EventID"]);
    
    calendars.forEach(function(calendar) {
      var events = calendar.getEvents(startTimeBas, endTime);
      
      events.forEach(function(event) {
        sheet.appendRow([
          calendar.getName(), // Имя календаря
          event.getTitle(), 
          event.getStartTime(), 
          event.getEndTime(), 
          event.getDescription(),
          event.getColor(),
          event.getId()
        ]);
      });
    });
  }


