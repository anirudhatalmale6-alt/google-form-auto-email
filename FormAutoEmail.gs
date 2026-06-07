// ============================================================
// GOOGLE FORM AUTO-EMAIL SCRIPT
// ============================================================
// This script does two things when someone submits your form:
//   1. Sends YOU a copy of their full submission
//   2. Sends THEM either a "correct" or "wrong" reply
//
// SETUP: See the instructions document for step-by-step setup.
// ============================================================


// ============================
// SETTINGS - CHANGE THESE
// ============================

// Your email address (where you get notified of each submission)
var ADMIN_EMAIL = "xpertusva@gmail.com";

// Subject line for the email YOU receive
var ADMIN_SUBJECT = "New Form Submission Received";

// The column header in your Sheet that contains the user's answer
// (Open your Sheet, look at Row 1, find the column with the answers)
var ANSWER_COLUMN = "Answer";

// The correct answer value (what the cell should contain to count as "right")
var CORRECT_ANSWER = "B";

// Subject line for the auto-reply the USER receives
var REPLY_SUBJECT_CORRECT = "Your Response - Well Done!";
var REPLY_SUBJECT_WRONG   = "Your Response - Thank You";


// ============================
// EMAIL TEMPLATES
// ============================
// Edit the text between the backticks (`) to change what users receive.
// You can use multiple lines. Keep the backticks at start and end.

var TEMPLATE_CORRECT = `Thanks for participating.

Well done! You got it right.

The correct answer is "B" - you should "switch" to the "Vanishing Edge"

This simple challenge was to demonstrate how often we arrive at the wrong decision. Of greater concern is that we are not aware that we got it wrong.

I do have a question: Did you guess?

Do you really understand why "SWITCH"?

Let me know, if you want to discuss know why I say it was the wrong decision, Happy to explain.

If you wish to learn more about who I am, and what I do, - study this 1-page document, and follow the links.

Thanks.
Dr. Errol: errolw@xpertus.com`;


var TEMPLATE_WRONG = `Thanks for participating.

The correct answer is "B" - you should "switch" to the "Vanishing Edge"

This simple challenge was to demonstrate how often we arrive at the wrong decision. Of greater concern is that we are not aware that we got it wrong.

Let me know, if you want to discuss know why I say it was the wrong decision, Happy to explain.

If you wish to learn more about who I am, and what I do, - study this 1-page document, and follow the links.

Thanks.
Dr. Errol: errolw@xpertus.com`;


// ============================================================
// MAIN FUNCTION - runs automatically on each form submission
// ============================================================

function onFormSubmit(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var lastRow = sheet.getLastRow();
    var rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Build a summary of the full submission for the admin email
    var submissionSummary = "";
    var userEmail = "";
    var userAnswer = "";

    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].toString().trim();
      var value = rowData[i].toString().trim();

      submissionSummary += header + ": " + value + "\n";

      // Find the user's email (Google Forms stores it as "Email Address" or "Email address")
      if (header.toLowerCase().indexOf("email") !== -1) {
        userEmail = value;
      }

      // Find the user's answer
      if (header.toLowerCase() === ANSWER_COLUMN.toLowerCase()) {
        userAnswer = value;
      }
    }

    // --- STEP 1: Send admin notification ---
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: ADMIN_SUBJECT,
      body: "New submission received:\n\n" + submissionSummary
    });

    // --- STEP 2: Send auto-reply to user ---
    if (userEmail) {
      // Check if the answer contains the correct value
      var isCorrect = userAnswer.toUpperCase().indexOf(CORRECT_ANSWER.toUpperCase()) !== -1;

      var replySubject = isCorrect ? REPLY_SUBJECT_CORRECT : REPLY_SUBJECT_WRONG;
      var replyBody = isCorrect ? TEMPLATE_CORRECT : TEMPLATE_WRONG;

      MailApp.sendEmail({
        to: userEmail,
        subject: replySubject,
        body: replyBody
      });
    }

  } catch (error) {
    // If something goes wrong, email the admin about the error
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "Form Script Error",
      body: "The auto-email script encountered an error:\n\n" + error.toString()
    });
  }
}


// ============================================================
// ONE-TIME SETUP FUNCTION - Run this once to create the trigger
// ============================================================

function createTrigger() {
  // Remove any old triggers first
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // Create new trigger that fires on form submission
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  Logger.log("Trigger created successfully! The script will now run on each form submission.");
}
