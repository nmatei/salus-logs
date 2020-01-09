/**
 * Run collectLogs every {x} hours 
 */
function collectLogs() {
  var logs = getLogsFromMail();
  if (logs.length) {
    //var sheet = SpreadsheetApp.getActiveSheet();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("logs");
    
    logs.forEach(function(log) {
      //sheet.insertRowAfter(1);
      sheet.insertRowBefore(2);
      
      var insertAt = 2,
        last = sheet.getLastRow() + 1,
        subject = log.subject,
        room = subject.replace(/.*log-h-(\w*)-.*/gi, "$1"),
        status = subject.indexOf("-off") >= 0 ? "OFF" : "ON";

      sheet.getRange(insertAt, 1).setValue(room);
      sheet.getRange(insertAt, 2).setValue(status);
      sheet.getRange(insertAt, 3).setValue(log.date);

      if (status === "OFF") {
        // if the last record is also OFF
        //  don't show the calculate time - it shows incorrect start time.
        var lastOff = "=maxifs($C$3:$C" + last + ", $A$3:$A" + last + ", $A" + insertAt + ", $B$3:$B" + last + ', "OFF")';
        sheet.getRange(insertAt, 4).setValue(lastOff);
        var lastOFFTime = new Date(sheet.getRange(insertAt, 4).getDisplayValue());

        var valueFormula = "=maxifs($C$2:$C" + last + ", $A$2:$A" + last + ", $A" + insertAt + ", $B$2:$B" + last + ', "ON")';
        sheet.getRange(insertAt, 4).setValue(valueFormula);
        sheet.getRange(insertAt, 5).setValue("=C" + insertAt + "-D" + insertAt);

        var startTime = new Date(cleanFormula(sheet, insertAt, 4));
        if (lastOFFTime.getTime() > startTime.getTime()) {
          // missing ON (2 OFF records wihtout start between)
          sheet.getRange(insertAt, 5).setValue("00:00:00.000");
        } else {
          cleanFormula(sheet, insertAt, 5);
        }
      }
    });
  }
}

/**
 * @private
 */
function getLogsFromMail() {
  var logs = [];
  var query = "from:(support@sc-smarthome.io OR support@salusconnect.io) subject:(OneTouch rule log-h-) -in:trash"; // is:unread
  var threads = GmailApp.search(query);

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = messages.length - 1; j >= 0; j--) {
      var mail = messages[j];

      if (mail.isInTrash()) {
        // STOP
        j = -1;
      } else {
        var subject = mail.getSubject();

        logs.push({
          subject: subject,
          date: mail.getDate()
        });

        mail.markRead();
        mail.moveToTrash();
      }
    }
  }
  
  logs.sort(function(a, b) {
    return a.date.getTime() - b.date.getTime();
  });
  return logs;
}

/**
 * @private
 * clean formulas to improve display performance
 */
function cleanFormula(sheet, rowIdx, cellIdx) {
  var value = sheet.getRange(rowIdx, cellIdx).getDisplayValue();
  sheet.getRange(rowIdx, cellIdx).setValue(value);
  return value;
}