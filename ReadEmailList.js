/* 
-------------------------------------------------------
Read Email Messages
-------------------------------------------------------
  
 Read all the email messages inlcuding the total size of the message and attachments, and write to spreadsheet
 
 This is slower than just reading the messages above

 This script reads Emails by threads (aka conversations) rather than 
 just reading through each individual email. Emails are therefore sorted 
 by the date of the initial email in the thread. The number of emails 
 returned in each thread is dependent on the number of emails (messages) in 
 a conversation, so each Block (threadBlockSize) of threads that is read, 
 may retreive more or less emails than the previous block
  
 If the script times out or the API limit is reached, around 25k,
 the script can be simply be rerun and it will continue from at the last line.

 Updates
 27 Oct 2024  RJ    Create EmailList worksheet if it doesn't exist. Force update of sheet after every block
 04 Nov 2024  RJ    Refactored to reduce code. Added Gmail Search column
*/

function readEmailMessages() {
  var logging = false;       // Turn off logging for faster processing
  var threadPage = 0;        // 0 is first page
  var threadBlockSize = 50;  // Number of threads to read before writing to the spreadsheet
  var newThread = "";        // Flag new a new Thread has started
  // Use the value in this field to search for the related email in Gmail  
  var gmailSearch = '=INDIRECT("F"&ROW()) & " after:" & TEXT(INDIRECT("E"&ROW())-1,"YYYY/MM/DD") & " before:" & TEXT(INDIRECT("E"&ROW())+1,"YYYY/MM/DD")';

  // Get the EmailList sheet or create a new one
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmailList");
  if (sheet == null) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("EmailList");
  }
  Logger.log("Started  ---------------------");
 
  // Restart from the last line or print the header if there are no lines
  var line = sheet.getLastRow();
  if (line > 1 ) {
    var data = sheet.getRange(line, 1, 1, sheet.getLastColumn()).getValues();  
    threadPage = data[0][0] +1; // Restart after last page
  } else {
    var header = [["threadPage","Thread", "#","Sender","Date","Subject", "Size (bytes)","New Thread", "Gmail Search"]];
    sheet.getRange(1, 1, 1, header[0].length).setValues(header).setFontWeight("bold");
    SpreadsheetApp.flush(); // Force write of the last Invoice values to the sheet 
  }
  
  // Read blocks of threads i.e. email + replies
  for (;; threadPage++) {
    var threads = GmailApp.search('-in:trash size:1M', threadPage*threadBlockSize, threadBlockSize)
    if (threads.length == 0) { break; } // No more threads
    if (logging) {Logger.log("threadPage:" + threadPage + " Threads:" + threads.length + " Line: " + line); }
  
    var messageLines = [];
    // Loop through the conversation threads in the block
    for (var threadCount = 0; threadCount < threads.length; threadCount++) {  
      var messages = threads[threadCount].getMessages();
      // Flag that a new conversation threads has started
      newThread = "***";
      if (logging) { Logger.log("*   threadPage:" + threadPage + " Thread:" + threadCount + " Line: " + line); }
  
      // Loop through the emails in a conversation thread
      for (var i = 0; i < messages.length; i++) {
        var messagesDetails = [threadPage, threadCount, line++, messages[i].getFrom(), messages[i].getDate(), messages[i].getSubject(), messages[i].getRawContent().length, newThread, gmailSearch];
        newThread = "";
        messageLines.push(messagesDetails);
      }
    }
    Logger.log("Print Lines --------------------- threadPage: " + threadPage + " Line: " + line);
    sheet.getRange(sheet.getLastRow()+1, 1, messageLines.length, messageLines[0].length).setValues(messageLines); 
    SpreadsheetApp.flush(); // Force write of the last Invoice values to the sheet 
  }
  Logger.log("Completed  ---------------------");
}
