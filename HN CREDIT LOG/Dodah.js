// function sendEmailOnThursday() {
//   // Hardcoded recipient, subject, and content
//   var recipient = "example@example.com";
//   var emailSubject = "Your Subject Here";
//   var emailContent = "Dear recipient,<br><br>This is the content of the email.<br><br>Sincerely,<br>Your Name";
  
//   // Get today's date
//   var today = new Date();
  
//   // Check if today is Thursday and time is 4:30 AM
//   if (today.getDay() == 4 && today.getHours() == 4 && today.getMinutes() == 30) {
//     // Send the email
//     MailApp.sendEmail({
//       to: recipient,
//       subject: emailSubject,
//       htmlBody: emailContent
//     });
//   }
// }

function sendEmailOnSaturday() {
  // Hardcoded recipient, subject, and content
  var recipient = "patrick.kell@momsorganicmarket.com";
  var emailSubject = "Your Subject Here";
  var emailContent = "Dear recipient,<br><br>This is the content of the email.<br><br>Sincerely,<br>Your Name";
  
  // Get today's date
  var today = new Date();
  
  // Check if today is Saturday and time is within the 30-minute window
  if (today.getDay() == 6 && today.getHours() == 14 && today.getMinutes() >= 30 && today.getMinutes() < 60) {
    // Send the email
    MailApp.sendEmail({
      to: recipient,
      subject: emailSubject,
      htmlBody: emailContent
    });
  }
}
