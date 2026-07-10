const DEMO_FEEDBACK_REQUEST_HEADERS = [
  "Feedback Token", "Booking Reference", "Client Email", "Client Name",
  "Event Date", "Event Type", "Request Due", "Request Sent", "Request Sent Date",
  "Opened", "Opened Date", "Completed", "Completed Date", "Booking Snapshot JSON",
  "Site ID"
];

const DEMO_FEEDBACK_RESPONSE_HEADERS = [
  "Feedback ID", "Booking Reference", "Submitted At", "Overall Satisfaction",
  "Food Quality", "Presentation", "Delivery Timing", "Ease Of Booking", "NPS",
  "What Went Well", "Improvements", "Additional Comments", "Contact Requested",
  "Preferred Contact Details", "Feedback JSON", "Site ID"
];

const DEMO_FEEDBACK_ITEM_HEADERS = [
  "Feedback ID", "Booking Reference", "Item ID", "Item Name", "Quantity",
  "Rating", "Comments", "Site ID"
];

function sendDemoFeedbackRequest_(dashboardBooking) {
  try {
    if (!SITE_CONFIG.feedback.enabled) return { sent: false, reason: "Feedback is disabled." };
    const settings = getPlatformSettings_();
    const recipient = normaliseDemoFeedbackRecipient_(settings.DEMO_FEEDBACK_RECIPIENT);
    if (!recipient) return { sent: false, reason: "No feedback recipient configured." };

    const webAppUrl = String(settings.FEEDBACK_WEB_APP_URL || SITE_CONFIG.feedback.webAppUrl || "").trim();
    if (!/^https:\/\/\S+\/exec(?:\?|$)/i.test(webAppUrl)) {
      return { sent: false, reason: "Feedback web app URL is missing or invalid." };
    }

    const spreadsheet = getDashboardSpreadsheet_();
    const requestSheet = getOrCreateSheet_(spreadsheet, "Feedback Requests", DEMO_FEEDBACK_REQUEST_HEADERS);
    getOrCreateSheet_(spreadsheet, "Feedback Responses", DEMO_FEEDBACK_RESPONSE_HEADERS);
    getOrCreateSheet_(spreadsheet, "Feedback Item Ratings", DEMO_FEEDBACK_ITEM_HEADERS);
    const token = Utilities.base64EncodeWebSafe(Utilities.getUuid() + Utilities.getUuid()).replace(/=+$/g, "");
    const now = new Date();
    requestSheet.appendRow([
      token,
      dashboardBooking.bookingId,
      dashboardBooking.hostEmail || recipient,
      dashboardBooking.hostName || "",
      dashboardBooking.eventDate || "",
      dashboardBooking.serviceType || "",
      now,
      true,
      now,
      false,
      "",
      false,
      "",
      JSON.stringify(dashboardBooking),
      SITE_CONFIG.siteId
    ]);

    sendDemoFeedbackEmail_(dashboardBooking, recipient, buildDemoFeedbackLink_(webAppUrl, token));
    return { sent: true, recipient: recipient };
  } catch (error) {
    console.error("Feedback request failed: " + error.message);
    return { sent: false, reason: error.message };
  }
}

function normaliseDemoFeedbackRecipient_(value) {
  const email = String(value || SITE_CONFIG.feedback.recipient || "").trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : "";
}

function buildDemoFeedbackLink_(webAppUrl, token) {
  return webAppUrl + (webAppUrl.indexOf("?") === -1 ? "?" : "&") +
    "token=" + encodeURIComponent(token);
}

function sendDemoFeedbackEmail_(booking, recipient, link) {
  const subject = "FIKA feedback link | " + booking.bookingId;
  const body = [
    "A FIKA hospitality booking has been submitted.",
    "",
    "Booking reference: " + booking.bookingId,
    "Client: " + (booking.clientCompany || ""),
    "Event date: " + (booking.eventDate || ""),
    "",
    "Open the feedback form:",
    link
  ].join("\n");
  const html = [
    '<div style="font-family:Gilroy,Arial,sans-serif;color:#280F8C;max-width:620px">',
    '<div style="background:#4F34C7;color:#fff;padding:24px 28px">',
    '<div style="font-size:12px;font-weight:bold;letter-spacing:1.3px;text-transform:uppercase">FIKA hospitality</div>',
    '<h1 style="margin:8px 0 0;font-size:26px">Feedback link ready</h1>',
    '</div>',
    '<div style="padding:26px 28px;border:1px solid #d9d4f4;border-top:0;background:#fff">',
    '<p>A FIKA booking has been submitted and is ready for feedback.</p>',
    '<p><strong>Reference:</strong> ' + escapeDemoFeedbackHtml_(booking.bookingId) + '</p>',
    '<p><strong>Client:</strong> ' + escapeDemoFeedbackHtml_(booking.clientCompany || '') + '</p>',
    '<p><strong>Date:</strong> ' + escapeDemoFeedbackHtml_(booking.eventDate || '') + '</p>',
    '<p style="margin:24px 0"><a href="' + escapeDemoFeedbackHtml_(link) + '" style="display:inline-block;background:#4DF7C2;color:#280F8C;padding:13px 18px;text-decoration:none;font-weight:bold">Leave feedback</a></p>',
    '</div></div>'
  ].join("");
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body,
    htmlBody: html,
    name: SITE_CONFIG.clientFacingName
  });
}

function escapeDemoFeedbackHtml_(value) {
  return String(value === null || value === undefined ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
