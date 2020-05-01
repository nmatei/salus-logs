/**
 * Run collectLogs every {x} hours 
 */
function collectLogs() {
  var roomMatch = /.*log-h-(.*)-([on]*[off]*).*/i;
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
        match = subject.match(roomMatch),
        room = match[1],
        status = match[2].toUpperCase();

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

        var startTimeStr = cleanFormula(sheet, insertAt, 4);
        var startTime = new Date(startTimeStr);
        if (lastOFFTime.getTime() > startTime.getTime()) {
          // missing ON (2 OFF records wihtout start between)
          //  could not find start time
          //  (don't print 12/30/1899 0:00:00, but insert start time)
          sheet.getRange(insertAt, 4).setValue(log.date);
          sheet.getRange(insertAt, 5).setValue("00:00:00.000");
        } else {
          cleanFormula(sheet, insertAt, 5);
          cleanupONRow(sheet, room, insertAt + 1, startTime);
        }
      }
    });
  }
}

/**
 * @private
 */
function cleanupONRow(sheet, room, startAt, startTime) {
  var index = findRowIndex(sheet, room, "ON", startAt, startTime);
  if (index != -1) {
    //Logger.log('delteRow', index);
    sheet.deleteRow(index); 
    return index;
  }
  return -1;
}

/**
 * @private
 */
function findRowIndex(sheet, room, status, startAt, startTime) {
  var time = startTime ? startTime.getTime() : false;
  for (var i = startAt; i < 100; i++) {
    var range = sheet.getRange(i, 1, 1, 3); // get one row
    var row = range.getValues()[0];
    if (row[0] == room && row[1] == status) {
      if (!time || row[2].getTime() == time) {
        //Logger.log("found startAt", i, room, row);
        return i;
      } else {
        //Logger.log("found startAt, wrong time", i, room, row);
      }
    }
  }
  return -1;
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

/**
 * @private
 */
function cleanupRoom(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("logs");
  //cleanupONRow(sheet, "living", 5, new Date("2/15/2020 20:03:21"));
  var index = 3;
  do {
    index = cleanupONRow(sheet, "living", index);
  } while (index != -1);
}