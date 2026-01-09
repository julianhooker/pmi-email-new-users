var thoughtspotMessageTitle = "New Members (new 2025) Liveboard update";
var pastNotificationsSpreadsheetFileId = "1tt3kzvSyhi4ZbCymHy3Imay5ki3rMtjs-8Xbze7J0JE" // Emails/New Member Emails/Sent - New Member Email
var newMemberMessageDocumentId = "1SH6CrsOH5CE6H0rfZJucCdBg1AA5aP-gTwsl-nn9vvw";
var newMemberMessageDocumentAttachmentId = "1p0PFYeb1RNm-fA5qG3gRTk8QHzKaNA9r";

var joinDatePosition = 1;
var firstNamePosition = 2;
var lastNamePosition = 3;
var emailPosition = 5;

var peopleWeEmailed = [];
var peopleNotNotified = [];

function myFunction() {
  Logger.log ("Starting Execution");

  var threads = GmailApp.search('subject:"' + thoughtspotMessageTitle + '"');

  Logger.log("Number of messages found: " + threads.length)
 

  var emailMessage = DocumentApp.openById(newMemberMessageDocumentId);
  var attachment = DriveApp.getFileById(newMemberMessageDocumentAttachmentId);
  var reportEmailFound = false;

  // Open the "Sent - New Member Email" spreadsheet once and cache its data to avoid repeated opens
  var pastNotificationsSheet = SpreadsheetApp.openById(pastNotificationsSpreadsheetFileId);
  var pastData = pastNotificationsSheet.getDataRange().getValues();

  // if there was an email from ThoughtSpot for the new members, we will process it
  for (var t = 0; t < threads.length; t++) {
    reportEmailFound = true;

    var messages = threads[t].getMessages();

    // Iterate messages
    for (var m = 0; m < messages.length; m++) {
      var attachments = messages[m].getAttachments();

      // Iterate attachments
      for (var a = 0; a < attachments.length; a++) {
        var dataExtract = Utilities.parseCsv(attachments[a].getDataAsString());

        // Just checking to see if this is a data extract from ThoughtSpot. If not, skip this attachment
        if (!dataExtract || !dataExtract[0] || dataExtract[0][0].indexOf("Data extract produced by") == -1) continue;

        // Find the start of the actual data (skip header). The original logic mutated loop counters and could run past the array bounds.
        var startIndex = 0;
        while (startIndex < dataExtract.length && dataExtract[startIndex][0] !== '') {
          startIndex++;
        }
        // Move past the blank line and the column titles row if possible
        startIndex = Math.min(dataExtract.length, startIndex + 2);

        for (var r = startIndex; r < dataExtract.length; r++) {
          var row = dataExtract[r];
          if (!row || row.length <= emailPosition) continue;

          var email = row[emailPosition];
          if (!email || email.toString().trim() === '') continue;

          if (memberNotifiedPreviously(email, pastData)) {
            peopleNotNotified.push(email);
          } else {
            var body = emailMessage.getText().replace("<first name>", row[firstNamePosition]);

            // Send email to the new member
            GmailApp.sendEmail(email, 'Welcome to the PMI West Texas Chapter!', body, {
                attachments: [attachment],
                htmlBody: body,
                name: 'PMI West Texas'
              });

            var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone ? Session.getScriptTimeZone() : "GMT+1", "MM/dd/yyyy");
            pastNotificationsSheet.appendRow([row[firstNamePosition], row[lastNamePosition], row[emailPosition], row[joinDatePosition], formattedDate]);

            // Keep in-memory summary up-to-date so subsequent checks in this run see the new entry
            pastData.push([row[firstNamePosition], row[lastNamePosition], row[emailPosition], row[joinDatePosition], formattedDate]);

            peopleWeEmailed.push(email);
          }
        }
      }
    }
  }

  Logger.log("reportEmailFound " + reportEmailFound);

  if (reportEmailFound) {
    var notificationMessage = "Just to let you know, a report from ThoughtSpot was found to process today <br /><br />";

    if (peopleWeEmailed.length != 0) {
      // notify VP membership who was emailed and who was not. 
      notificationMessage += "New members were found to notify. Details are below. <br />";

      notificationMessage += "Here are the people who we sent the notification to just now: ";
      notificationMessage += "<ul>";
      peopleWeEmailed.forEach((value) => {
        notificationMessage += "<li>" + value + "</li>";
      })
      notificationMessage += "</ul>";
    } else {
      notificationMessage += "No new people in the report. <br /><br />";
    }

    if (peopleNotNotified.length != 0) {
      notificationMessage += "People not notified: ";
      notificationMessage += "<ul>";
      peopleNotNotified.forEach((value) => {
        notificationMessage += "<li>" + value + "</li>";
      })
      notificationMessage += "</ul>";
    } else {
      notificationMessage += "There were no people who have already been notified in the report. <br /><br />";
    }

    Logger.log("Sending report email");
    GmailApp.sendEmail('membership@pmiwtx.org', 'New Member Notification Report', notificationMessage, {
      htmlBody: notificationMessage,
      name: 'PMI West Texas'
    })

    Logger.log("Deleting original email");
    // now delete that email to keep the inbox clean
    threads.forEach((value) => {
      //value.moveToTrash();
    })
  }
}

function memberNotifiedPreviously(email, pastData) {
    // pastData: optional cached 2D array of rows from the past notifications sheet
    var data = pastData;
    if (!data) {
      var pastNotificationsSheet = SpreadsheetApp.openById(pastNotificationsSpreadsheetFileId);
      data = pastNotificationsSheet.getDataRange().getValues();
    }

    if (!email) return false;
    var target = email.toString().trim().toUpperCase();

    for (var i = 0; i < data.length; i++) {
      var rowEmail = (data[i] && data[i][2]) ? data[i][2].toString().trim().toUpperCase() : '';
      if (target === rowEmail) return true; // found the email address in the sheet
    }

    return false; // did not find the email address in the sheet
}