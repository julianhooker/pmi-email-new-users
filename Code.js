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

  // if there was an email from ThoughtSpot for the new members, we will process it
  for (var i = 0; i < threads.length; i++) {
    reportEmailFound = true;

    var messages = threads[i].getMessages();

    // Let's get the messages
    for (var i = 0; i < messages.length; i++) {
      var attachments = messages[i].getAttachments();

      // Let's get the attachment from the message
      for (var i = 0; i < attachments.length; i++) {
        var dataExtract = Utilities.parseCsv(attachments[i].getDataAsString());

        // Just checking to see if this is a data extract from ThoughtSpot. If not, stop executing
        if (dataExtract[0][0].indexOf("Data extract produced by") == -1) break;

        // Open the "Sent - New Member Email" spreadsheet. This keeps track of who has already been send the new memeber email
        var pastNotificationsSheet = SpreadsheetApp.openById(pastNotificationsSpreadsheetFileId);
        var data = pastNotificationsSheet.getDataRange().getValues();

        var processingHeader = true;
        for (var i = 0; i < dataExtract.length; i++) {
          while (processingHeader) {
            if (dataExtract[i][0] == '') {
              // found the blank end of the header
              i++; // Move past the blank line
              i++; // Move past the row with the column titles in it

              processingHeader = false;
            } else {
              i++
            }
          }

          // Logger.log('First Name: ' + dataExtract[i][firstNamePosition]);
          // Logger.log('Last Name: ' + dataExtract[i][lastNamePosition]);
          // Logger.log('Email Address: ' + dataExtract[i][emailPosition]);
          // Logger.log('Joined: ' + dataExtract[i][joinDatePosition]);

          if (memberNotifiedPreviously(dataExtract[i][emailPosition])) {
            //Logger.log("Email address " + dataExtract[i][emailPosition] + " is in the spreadsheet");

            peopleNotNotified.push(dataExtract[i][emailPosition]);
          } else {
              var body = emailMessage.getText().replace("<first name>", dataExtract[i][firstNamePosition]);

              // This will email to the new members

              //Logger.log("Would have emailed " + dataExtract[i][emailPosition])
              GmailApp.sendEmail(dataExtract[i][emailPosition], 'Welcome to the PMI West Texas Chapter!', body, {
                  attachments: [attachment],
                  htmlBody: body,
                  name: 'PMI West Texas'
                });

              pastNotificationsSheet
              .appendRow([dataExtract[i][firstNamePosition], dataExtract[i][lastNamePosition], dataExtract[i][emailPosition], dataExtract[i][joinDatePosition], Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy")]);

              peopleWeEmailed.push(dataExtract[i][emailPosition]);
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

function memberNotifiedPreviously(email, pastNotificationsSheet) {
    var pastNotificationsSheet = SpreadsheetApp.openById(pastNotificationsSpreadsheetFileId);
    var data = pastNotificationsSheet.getDataRange().getValues();

    for (var i = 0; i < data.length; i++) {
      if (email.trim().toUpperCase() == data[i][2].trim().toUpperCase()) return true; // found the email address in the sheet
    }

    return false; // did not find the email address in the sheet
}