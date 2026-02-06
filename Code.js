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
// Set to true to redirect all outgoing member emails to `TEST_RECIPIENT` for safe testing
var TEST_MODE = false;
var TEST_RECIPIENT = 'membership@pmiwtx.org';

// Conditional logger: only logs when TEST_MODE is true
function log(message) {
  if (TEST_MODE) Logger.log(message);
}

function myFunction() {
  log("Starting Execution");

  var threads = GmailApp.search('in:inbox subject:"' + thoughtspotMessageTitle + '"');

  log("Number of messages found: " + threads.length)
 

  var emailMessage = DocumentApp.openById(newMemberMessageDocumentId);
  var attachment = DriveApp.getFileById(newMemberMessageDocumentAttachmentId);


  var reportEmailFound = false;

  // if there was an email from ThoughtSpot for the new members, we will process it
  for (var tIdx = 0; tIdx < threads.length; tIdx++) {
    reportEmailFound = true;
    log('Processing thread ' + tIdx + ' of ' + threads.length);

    var messages = threads[tIdx].getMessages();

    // Let's get the messages
    for (var mIdx = 0; mIdx < messages.length; mIdx++) {
      log(' Processing message ' + mIdx + ' of ' + messages.length + ' in thread ' + tIdx);
      var attachments = messages[mIdx].getAttachments();

      // Let's get the attachment from the message
      for (var aIdx = 0; aIdx < attachments.length; aIdx++) {
        var dataExtract = Utilities.parseCsv(attachments[aIdx].getDataAsString());
        log('  Found attachment ' + aIdx + ' with ' + dataExtract.length + ' rows (including header)');

        // Just checking to see if this is a data extract from ThoughtSpot. If not, stop executing
        if (dataExtract[0][0].indexOf("Data extract produced by") == -1) break;

        // Open the "Sent - New Member Email" spreadsheet. This keeps track of who has already been send the new memeber email
        var pastNotificationsSheet = SpreadsheetApp.openById(pastNotificationsSpreadsheetFileId);
        var data = pastNotificationsSheet.getDataRange().getValues();

        var processingHeader = true;
        for (var r = 0; r < dataExtract.length; r++) {
          while (processingHeader) {
            if (dataExtract[r][0] == '') {
              // found the blank end of the header
              r++; // Move past the blank line
              r++; // Move past the row with the column titles in it

              processingHeader = false;
            } else {
              r++
            }
          }

          // Logger.log('First Name: ' + dataExtract[r][firstNamePosition]);
          // Logger.log('Last Name: ' + dataExtract[r][lastNamePosition]);
          // Logger.log('Email Address: ' + dataExtract[r][emailPosition]);
          // Logger.log('Joined: ' + dataExtract[r][joinDatePosition]);

          log('   Processing row ' + r + ': email=' + dataExtract[r][emailPosition]);
          if (memberNotifiedPreviously(dataExtract[r][emailPosition])) {
            log('    Email already notified previously: ' + dataExtract[r][emailPosition]);
            peopleNotNotified.push(dataExtract[r][emailPosition]);
          } else {
              var body = emailMessage.getText().replace("<first name>", dataExtract[r][firstNamePosition]);

              // This will email to the new members

              //Logger.log("Would have emailed " + dataExtract[r][emailPosition])
              var recipient = TEST_MODE ? TEST_RECIPIENT : dataExtract[r][emailPosition];
              GmailApp.sendEmail(recipient, 'Welcome to the PMI West Texas Chapter!', body, {
                  attachments: [attachment],
                  htmlBody: body,
                  name: 'PMI West Texas'
                });

              pastNotificationsSheet
              .appendRow([dataExtract[r][firstNamePosition], dataExtract[r][lastNamePosition], dataExtract[r][emailPosition], dataExtract[r][joinDatePosition], Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy")]);

              peopleWeEmailed.push(dataExtract[r][emailPosition]);
          }
        }
      }
    } 
  }

  log("reportEmailFound " + reportEmailFound);

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

    log("Sending report email");
    GmailApp.sendEmail('membership@pmiwtx.org', 'New Member Notification Report', notificationMessage, {
      htmlBody: notificationMessage,
      name: 'PMI West Texas'
    })

    log("Deleting original email");
    // now delete that email to keep the inbox clean
    threads.forEach((value) => {
      value.moveToTrash();
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
