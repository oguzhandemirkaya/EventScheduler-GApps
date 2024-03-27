function createEventFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sayfa1");
  var values = sheet.getDataRange().getValues();
  var calendarId = 'CALENDARID'; // Enter your Calendar ID here
  var calendar = CalendarApp.getCalendarById(calendarId);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[7] !== "Added") {
      var eventTitle = row[0];
      var eventDescription = 'Phone: ' + row[1] + ', E-mail: ' + row[3] + ', Country: ' + row[2];
      var eventDate = new Date(row[6]);
      var endDate = new Date(eventDate);
      endDate.setHours(eventDate.getHours() + 1);

      calendar.createEvent(eventTitle, eventDate, endDate, {description: eventDescription});
      sheet.getRange(i + 1, 8).setValue("Added");
      sendInstantEmail(row[3], row[0], i);
    }
  }
}

function sendInstantEmail(emailAddress, personName, rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sayfa1");
  var emailSentCell = sheet.getRange(rowIndex + 1, 9);

  if (!emailSentCell.getValue()) {
    var emailSubject = "NAME"; // Brand name 
    var emailBody = "Hello, " + personName + ", your registration has been received";
    GmailApp.sendEmail(emailAddress, emailSubject, emailBody);
    emailSentCell.setValue("Sent");
  }
}


// Daily Check Function
function dailyCheck() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sayfa1");
  var values = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var emailAddress = row[3];
    var personName = row[0];
    var eventDate = new Date(row[6]);

    checkAndSendEmailsForDate(emailAddress, personName, eventDate, i);
  }
}
//Sending mail function by adding days to dates
function checkAndSendEmailsForDate(emailAddress, personName, eventDate, rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("page1");
  var now = new Date();
  now.setHours(0, 0, 0, 0);

  var timeDifferences = [7, 14, 30, 90, 180, 240, 365];
  var emailSubjects = [
    "First Week Check Reminder",
    "Second Week Check Reminder",
    "First Month Check Reminder",
    "Third Month Check Reminder",
    "Sixth Month Check Reminder",
    "Eighth Month Check Reminder",
    "One Year Check Reminder"
  ];

  var messages = [
    "We hope you're doing well. It's time for your first week check. Please remember to follow the instructions provided.",
    "Reminder for your second week check. Please ensure you have followed the instructions and provide any necessary feedback.",
    "It's already one month since your procedure. Please remember to check in with us as instructed.",
    "Three months have passed since your procedure. We would like to remind you to follow up according to the instructions provided.",
    "Six months have passed, and it's time for another check. Please follow the provided instructions for the check-up.",
    "Eight months have gone by. Please remember to check in with us as instructed.",
    "It's been a year since your procedure. Please remember to complete your one-year check as per the instructions."
  ];

  var emailSentColumns = [10, 11, 12, 13, 14, 15, 16];

  for (var i = 0; i < timeDifferences.length; i++) {
    var targetDate = new Date(eventDate);
    targetDate.setDate(eventDate.getDate() + timeDifferences[i]);
    targetDate.setHours(0, 0, 0, 0);
    var emailSentCell = sheet.getRange(rowIndex + 1, emailSentColumns[i]);

    if (isSameDay(targetDate, now) && !emailSentCell.getValue()) {
      var emailSubject = emailSubjects[i];
      var personalizedGreeting = "Hello, " + personName + ",\n\n";
      var emailBody = personalizedGreeting + messages[i];

      GmailApp.sendEmail(emailAddress, emailSubject, emailBody);
      emailSentCell.setValue("Sent");
    }
  }
}
//  Comparison of two dates
function isSameDay(date1, date2) {
  return date1.getDate() === date2.getDate() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getFullYear() === date2.getFullYear();
}
