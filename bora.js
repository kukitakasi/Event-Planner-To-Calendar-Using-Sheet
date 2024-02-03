function onOpen() {
 createMenuWithSubMenu();
}

function createMenuWithSubMenu() {
  SpreadsheetApp.getUi().createMenu("ε(´｡•᎑•`)っ 💕")
   .addItem("🧪 Get Result", "filterAndCopyData")
   .addSeparator()
   .addItem("📩 Set Calender", "setCalendarEvents")
   .addSeparator()
   .addToUi();
}

function sendBulkEmails() {
 const data = ["raushan@corenexis.com","sia@corenexis.com","your email goes here"] //you can also add more emails also you can give an array range by google sheet.


 // Subject and body of the email
 let subject = 'Hii, System Maintanance TO-DO is Here';
 let body = "Hi,\n\nI hope you are doing well.\n\nThis email is a reminder about your renewal plans of system maintainance.\n\nThanks for your attention,\n Do not reply to this automated email.\n\n Powered By kukitakasi :) \n\n ";




 // Loop through each email address and send the email
 for (var i = 0; i < data.length; i++) {
   var emailAddress = data[i];
  
   // Check if the email address is not empty
   if (emailAddress && emailAddress.trim() !== "") {
     sendEmail(emailAddress, subject, body+emailAddress);
   } else {
     Logger.log('Skipping empty email address at row ' + (i + 2));
   }
 }
}


function sendEmail(emailAddress, subject, body) {
 try {
   GmailApp.sendEmail(emailAddress, subject, body);
   Logger.log('Email sent to: ' + emailAddress);
 } catch (error) {
   Logger.log('Error sending email to ' + emailAddress + ': ' + error.toString());
 }
}


function setCalendarEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Expiry Alert");
  const range = sheet.getRange("A3:H200");
  const values = range.getValues();
  const tickBoxRange = sheet.getRange("H3:H200");
  const tickBoxValues = tickBoxRange.getValues();
  
 const calendar = CalendarApp.getCalendarById("your calendar id");
  
  for (let i = 0; i < values.length; i++) {
    const recordId = values[i][0]; // Record ID from column A
    const title = values[i][2]; // Title from column B
    const description = values[i][3]; // Description from column D
    const date = values[i][4]; // Date from column E
    const tickBox = tickBoxValues[i][0]; // Tick box value from column H
    
    if (recordId && title && description && date && tickBox === true) {
      // Parse the date string into a Date object
      const eventDate = new Date(Date.parse(date));
      
      // Set start and end dates for the event (30 minutes before and after)
      const startDate = new Date(eventDate.getTime() - 30 * 60 * 1000); // 30 minutes before
      const endDate = new Date(eventDate.getTime() + 30 * 60 * 1000); // 30 minutes after
      
      // Create the event with a 30-minute duration and add the record ID in the description
      const event = calendar.createEvent(title, startDate, endDate);
      event.setDescription("Record ID: " + recordId + "\n" + description);
      
      // Clear the tick box in column H for the processed row
      tickBoxRange.getCell(i + 1, 1).setValue(" ");
    }
  }
}
