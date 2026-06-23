function sendFeedbackRequestEmail_(context, booking, token) {
  const settings = getFeedbackSettings_(context);
  const baseUrl = String(settings.FEEDBACK_WEB_APP_URL || "").trim();
  if (!/^https:\/\/\S+\/exec(?:\?|$)/i.test(baseUrl)) {
    throw new Error("FEEDBACK_WEB_APP_URL is missing or is not a deployed /exec URL.");
  }
  const link = baseUrl + (baseUrl.indexOf("?") === -1 ? "?" : "&") +
    "token=" + encodeURIComponent(token);
  const subject = "How was your hospitality booking?";
  const body = [
    "Thank you for booking hospitality with FIKA.",
    "",
    "We'd love your feedback. It should only take a minute.",
    "",
    link
  ].join("\n");
  const html = [
    '<div style="font-family:Arial,sans-serif;color:#221874;max-width:600px">',
    '<div style="background:#3d21bf;color:#fff;padding:26px;border-radius:14px 14px 0 0">',
    '<div style="font-size:12px;font-weight:bold;letter-spacing:1.4px;text-transform:uppercase">FIKA hospitality</div>',
    '<h1 style="margin:8px 0 0;font-size:28px">How did we do?</h1>',
    '</div>',
    '<div style="padding:28px;border:1px solid #dedbea;border-top:0;border-radius:0 0 14px 14px">',
    '<p>Thank you for booking hospitality with FIKA.</p>',
    '<p>We would love your feedback. It should only take a minute.</p>',
    '<p style="margin:24px 0"><a href="' + feedbackEscapeHtml_(link) + '" style="display:inline-block;background:#75efb8;color:#241176;padding:13px 20px;border-radius:9px;text-decoration:none;font-weight:bold">Leave feedback</a></p>',
    '<p style="font-size:12px;color:#716c8e">Booking reference: ' + feedbackEscapeHtml_(booking.bookingId) + '</p>',
    '</div></div>'
  ].join("");
  MailApp.sendEmail({
    to: booking.hostEmail,
    subject: subject,
    body: body,
    htmlBody: html,
    name: context.site.clientFacingName
  });
}

function sendFeedbackRecoveryNotification_(context, feedback, booking) {
  const settings = getFeedbackSettings_(context);
  const recipients = parseFeedbackEmails_([
    settings.SITE_EMAIL_ADDRESS,
    settings.FOLLOW_UP_RECIPIENTS
  ].filter(Boolean).join(","));
  if (!recipients.length) return;
  MailApp.sendEmail({
    to: recipients.join(","),
    subject: "Low hospitality feedback requires follow-up | " +
      context.site.siteName + " | " + feedback.bookingReference,
    body: [
      "A client has requested contact after leaving a low rating.",
      "",
      "Booking: " + feedback.bookingReference,
      "Company: " + (booking.clientCompany || ""),
      "Overall score: " + feedback.overallSatisfaction + "/5",
      "Preferred contact: " + feedback.preferredContactDetails,
      "Improvements: " + feedback.improvements,
      "Additional comments: " + feedback.additionalComments
    ].join("\n")
  });
}

function parseFeedbackEmails_(value) {
  const seen = {};
  return String(value || "").split(/[\s,;]+/)
    .map(function(email) { return email.trim(); })
    .filter(function(email) {
      const key = email.toLowerCase();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || seen[key]) return false;
      seen[key] = true;
      return true;
    });
}

function feedbackEscapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
