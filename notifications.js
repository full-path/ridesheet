// Pull notification email(s) from Doc Props
function sendNotification(subject, body) {
  let defaultEmail = "nome@garnetconsultingpdx.com"
  MailApp.sendEmail(defaultEmail, subject, body)
}
