function sendDodahsDateCheckEmail() {
  // Hardcoded recipient, subject, and content
  var recipient = "patrick.kell@momsorganicmarket.com";
  // var recipient = "hn.grocery@momsorganicmarket.com, hn.ops@momsorganicmarket.com, hn.customerservice@momsorganicmarket.com";
  var emailSubject = "Dodah's Date Check";
  var emailContent = "<p>Good morning team,</p><p>Please check the <strong><em>expiration dates of all Dodah's products</em></strong> on the shelf today. This helps us get credit for expired items when the delivery arrives this afternoon.</p><p>Let me know if you have any questions.</p><p>Thank you,</p>";

  // Send the email
  MailApp.sendEmail({
      to: recipient,
      subject: emailSubject,
      htmlBody: emailContent
  });
}
