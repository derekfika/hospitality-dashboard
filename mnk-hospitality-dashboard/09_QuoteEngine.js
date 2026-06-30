function generateQuoteForRow(rowNumber) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");

  booking = ensureLineItemTimes_(booking);
  booking = validateBooking_(booking);

  if (booking.validationErrors.length > 0) {
    throw new Error("Cannot generate quote. Missing: " + booking.validationErrors.join(", "));
  }

  const folder = getQuoteFolderForBooking_(booking);
  const quoteName = makeQuoteName_(booking);

  let quoteFile;

  if (booking.quoteUrl) {
    const existingId = extractDriveIdFromUrl_(booking.quoteUrl);
    if (existingId) {
      quoteFile = DriveApp.getFileById(existingId);
    }
  }

  if (!quoteFile) {
    const templateId = getConfiguredValue_("QUOTE_TEMPLATE_DOC_ID", CONFIG.QUOTE_TEMPLATE_DOC_ID);
    const template = DriveApp.getFileById(templateId);
    quoteFile = template.makeCopy(quoteName, folder);
  } else {
    quoteFile.setName(quoteName);
  }

  const doc = DocumentApp.openById(quoteFile.getId());
  clearAndRefillQuoteDoc_(doc, booking);
  doc.saveAndClose();

  booking.quoteUrl = quoteFile.getUrl();
  booking.quoteCreatedAt = new Date();
  booking.status = CONFIG.STATUS.QUOTE_GENERATED;
  booking.quoteStale = false;
  booking.updatedAt = new Date();

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return {
    ok: true,
    quoteUrl: quoteFile.getUrl()
  };
}

function ensureLineItemTimes_(booking) {
  const defaultTime =
    booking.serviceTimes && booking.serviceTimes.length
      ? booking.serviceTimes[0]
      : "";

  booking.items = (booking.items || []).map(item => {
    if (!item.time) item.time = defaultTime;
    item.time = String(item.time || "").trim();
    return item;
  });

  return booking;
}

function clearAndRefillQuoteDoc_(doc, booking) {
  const body = doc.getBody();

  const replacements = {
    "<SITE>": booking.location || getConfiguredValue_("LOCATION_NAME", CONFIG.LOCATION_NAME || CONFIG.APP_NAME || ""),
    "<FLOOR>": booking.floor || "",
    "<DAY>": formatDayName_(booking.eventDate),
    "<PAX>": booking.pax || "",
    "<DATE>": formatUkDate_(booking.eventDate),
    "<SERVICE>": booking.serviceType || "",
    "<DELIVERYTIME>": (booking.serviceTimes || []).join(" / "),
    "<SERVICETIME>": (booking.serviceTimes || []).join(" / "),
    "<LOCATION>": booking.location || "",
    "<HOST>": booking.hostName || booking.hostEmail || "",
    "<INVOICEREFERENCE>": booking.invoiceReference || "",
    "<INVOICE_REFERENCE>": booking.invoiceReference || "",
    "<NOTES>": booking.notes || "",
    "<TOTAL_PRICE>": formatMoney_(booking.totalPrice),
    "<MGMTFEE>": formatMoney_(booking.mgmtFee),
    "<NETPRICE>": formatMoney_(booking.netPrice),
    "<VAT>": formatMoney_(booking.vat),
    "<GROSSPRICE>": formatMoney_(booking.grossPrice)
  };

  Object.keys(replacements).forEach(key => {
    body.replaceText(escapeRegex_(key), String(replacements[key] || ""));
  });

  styleQuoteNotes_(body, booking.notes || "");

  replaceQuoteOrderPlaceholder_(body, booking.items || []);
}

function styleQuoteNotes_(body, notesText) {
  if (!notesText) return;

  const cleanNotes = String(notesText).trim();
  if (!cleanNotes) return;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const p = child.asParagraph();
    const text = p.getText();

    if (text.includes(cleanNotes)) {

      p.setForegroundColor("#C05050");
      p.setBold(true);
      p.setFontSize(10);

      const textElement = p.editAsText();

      textElement.setBold(0, text.length - 1, true);
      textElement.setForegroundColor(0, text.length - 1, "#C05050");

      return;
    }
  }
}

function replaceQuoteOrderPlaceholder_(body, items) {
  let targetIndex = -1;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const text = child.asParagraph().getText();
    if (text.includes("<ORDER>")) {
      targetIndex = i;
      break;
    }
  }

  if (targetIndex === -1) return;

  body.removeChild(body.getChild(targetIndex));

  let insertIndex = targetIndex;
  let currentSection = "";

  items.forEach(item => {
    if (item.section && item.section !== currentSection) {
      currentSection = item.section;

      const section = body.insertParagraph(insertIndex++, currentSection.toUpperCase());
      section.setBold(true);
      section.setFontSize(11);
      section.setSpacingBefore(8);
      section.setSpacingAfter(2);
    }

    const line = body.insertParagraph(insertIndex++, "");
    line.setFontSize(10);
    line.setSpacingAfter(2);

    const time = item.time ? item.time + " - " : "";
    const qty = item.qty ? " x " + item.qty : "";

    line.appendText(time).setBold(false).setItalic(false).setFontSize(10);
    line.appendText(item.name || "").setBold(true).setItalic(false).setFontSize(10);
    line.appendText(qty).setBold(false).setItalic(false).setFontSize(10);

    const detailParts = [];
    if (item.detail) detailParts.push(item.detail);
    if (item.info) detailParts.push(item.info);
    if (item.comment) detailParts.push("Comment: " + item.comment);

    if (detailParts.length) {
      const detail = body.insertParagraph(insertIndex++, detailParts.join(" - "));
      detail.setFontSize(8);
      detail.setItalic(true);
      detail.setSpacingAfter(4);
    }
  });
}

function getQuoteFolderForBooking_(booking) {
  const rootFolderName = getConfiguredValue_("QUOTE_ROOT_FOLDER_NAME", CONFIG.QUOTE_ROOT_FOLDER_NAME || "Hospitality");
  const root = getOrCreateDriveFolder_(rootFolderName);

  const year = booking.eventDate
    ? String(booking.eventDate).slice(0, 4)
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");

  const month = booking.eventDate
    ? formatFolderMonth_(booking.eventDate)
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM MMMM");

  const yearFolder = getOrCreateChildFolder_(root, year);
  return getOrCreateChildFolder_(yearFolder, month);
}

function getOrCreateDriveFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function makeQuoteName_(booking) {
  const date = booking.eventDate || "No Date";
  const company = booking.clientCompany || "Unknown Company";
  return `Quote - ${company} - ${date}`;
}

function formatMoney_(n) {
  return "GBP " + Number(n || 0).toFixed(2);
}

function formatUkDate_(isoDate) {
  if (!isoDate) return "";
  const parts = String(isoDate).split("-");
  if (parts.length !== 3) return isoDate;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function formatDayName_(isoDate) {
  if (!isoDate) return "";
  const parts = String(isoDate).split("-");
  if (parts.length !== 3) return "";
  const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "EEEE");
}

function formatFolderMonth_(isoDate) {
  const parts = String(isoDate).split("-");
  const d = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "MM MMMM");
}

function extractDriveIdFromUrl_(url) {
  const m = String(url || "").match(/[-\w]{25,}/);
  return m ? m[0] : "";
}

function escapeRegex_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function printQuoteForRow(rowNumber) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);
  const siteName = getConfiguredValue_(
  "LOCATION_NAME",
  "FIKA Hospitality"
  );
  
  if (!booking) throw new Error("Could not read booking data.");
  if (!booking.quoteUrl) throw new Error("No quote has been generated yet.");

  const fileId = extractDriveIdFromUrl_(booking.quoteUrl);
  if (!fileId) throw new Error("Could not read quote file ID.");

  const pdfBlob = DriveApp
    .getFileById(fileId)
    .getBlob()
    .getAs(MimeType.PDF)
    .setName(makeQuoteName_(booking) + ".pdf");

  MailApp.sendEmail({
    to: getConfiguredValue_("PRINTER_EMAIL", CONFIG.PRINTER_EMAIL),
    subject: `${siteName} Quote`,
    body: `Auto-printed quote from ${siteName} Dashboard.`
});

  booking.quotePrintedAt = new Date();
  booking.updatedAt = new Date();

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return {
    ok: true
  };
}

function sendBookingConfirmationEmail_(booking) {
  const ccAddress = getConfiguredValue_("CALENDAR_ID", CONFIG.CALENDAR_ID || "");

  const subject =
    `FIKA Hospitality | Booking Confirmed | ${formatEmailDate_(booking.eventDate)}`;

  const htmlBody = buildConfirmationEmailHtml_(booking);
  const plainTextBody = stripHtml_(htmlBody);

  if (!booking.hostEmail) {
  throw new Error("Cannot send confirmation email. Host email is missing.");
  }

  GmailApp.sendEmail(
    booking.hostEmail,
    subject,
    plainTextBody,
    {
      htmlBody: htmlBody,
      cc: ccAddress
    }
  );

  return {
    sentTo: booking.hostEmail
  };
}

function buildConfirmationSubject_(booking) {
  return `FIKA Hospitality | Booking Confirmed | ${formatEmailDate_(booking.eventDate)}`;
}

function buildConfirmationEmailHtml_(booking) {
  return `
  <div style="font-family: Arial, sans-serif; color:#241F33; line-height:1.5;">
    <h2 style="color:#4F34C7; margin-bottom:8px;">Booking Confirmed</h2>

    <p>Hi there!,</p>

    <p>Thank you for your booking with <strong>FIKA Hospitality</strong>.</p>

    <p>
      We're pleased to confirm that your booking has been received and scheduled with our team.
      Our catering and logistics teams have been notified and will begin preparing your order.
    </p>

    <div style="margin:22px 0; padding:16px; border:1px solid #DDD8EA; border-radius:14px; background:#F4F0FF;">
      <p><strong>Booking Reference:</strong> ${escapeEmailHtml_(booking.bookingId)}</p>
      <p><strong>Service:</strong> ${escapeEmailHtml_(booking.serviceType || "Hospitality")}</p><p><strong>Date:</strong> ${escapeEmailHtml_(formatEmailDate_(booking.eventDate))}</p>
      <p><strong>Time:</strong> ${escapeEmailHtml_((booking.serviceTimes || []).join(", "))}</p>
      <p><strong>Location:</strong> ${escapeEmailHtml_(booking.location || "")}</p>
      <p><strong>Floor:</strong> ${escapeEmailHtml_(booking.floor || "")}</p>
      <p><strong>Guests:</strong> ${escapeEmailHtml_(booking.pax || "")}</p>
    </div>

    <p>Should any changes be required before the service date, please let us know as soon as possible and we'll do our best to accommodate them.</p>

    <p style="margin-top:24px;">
      We look forward to enhancing your hospitality experience.<br><br>
      <strong style="color:#4F34C7; font-size:18px;">Let's FIKA!</strong>
    </p>

    <p style="margin-top:24px;">
      Kind regards,<br>
      <strong>FIKA Hospitality</strong>
    </p>
  </div>
  `;
}

function escapeEmailHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function stripHtml_(html) {
  return String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<[^>]+>/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function formatEmailDate_(isoDate) {
  if (!isoDate) return "";

  const parts = String(isoDate).split("-");
  if (parts.length !== 3) return isoDate;

  const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    "dd MMMM yyyy"
  );
}

function cancelBookingForRow(rowNumber, options) {
  options = options || {};

  const sh = getDashboardSheet_();
  const map = getHeaderMap_();
  const calendarId = getConfiguredValue_("CALENDAR_ID", CONFIG.CALENDAR_ID || "primary");

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");

  const result = {
    ok: true,
    bookingId: booking.bookingId,
    emailSent: false,
    calendarRemoved: false
  };

  if (options.sendEmail) {
    sendBookingCancellationEmail_(booking);
    result.emailSent = true;
  }

  if (options.removeCalendar && booking.calendarEventId) {
    Calendar.Events.remove(
      calendarId,
      booking.calendarEventId,
      {
        sendUpdates: "all"
      }
    );

    result.calendarRemoved = true;

    booking.calendarRemovedAt = new Date();
    booking.calendarEventId = "";
    booking.calendarEventUrl = "";
    booking.calendarStale = false;
  }

  booking.status = CONFIG.STATUS.CANCELLED || "CANCELLED";
  booking.cancelledAt = new Date();
  booking.cancelledBy = Session.getActiveUser().getEmail();
  booking.cancellationEmailSentAt = result.emailSent ? new Date() : "";
  booking.updatedAt = new Date();

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return result;
}

function sendBookingCancellationEmail_(booking) {
  const to = booking.hostEmail;
  const cc = getConfiguredValue_("CALENDAR_ID", CONFIG.CALENDAR_ID || "");

  if (!to) throw new Error("Cannot send cancellation email. Host email is missing.");

  const subject =
    `FIKA Hospitality | Booking Cancelled | ${formatEmailDate_(booking.eventDate)}`;

  const htmlBody = buildCancellationEmailHtml_(booking);

  GmailApp.sendEmail(to, subject, stripHtml_(htmlBody), {
    htmlBody,
    name: "FIKA Hospitality",
    cc
  });

  return { sentTo: to };
}

function buildCancellationEmailHtml_(booking) {
  return `
  <div style="font-family: Arial, sans-serif; color:#241F33; line-height:1.5; padding:24px;">
    <h2 style="color:#FF5C00; margin-bottom:8px;">Booking Cancelled</h2>

    <p>Hi there,</p>

    <p>
      This email is to confirm that the following FIKA Hospitality booking has been cancelled.
    </p>

    <div style="margin:22px 0; padding:18px; border:1px solid #DDD8EA; border-radius:14px; background:#FFF7F2;">
      <p><strong>Booking Reference:</strong> ${escapeEmailHtml_(booking.bookingId)}</p>
      <p><strong>Service:</strong> ${escapeEmailHtml_(booking.serviceType || "Hospitality")}</p>
      <p><strong>Date:</strong> ${escapeEmailHtml_(formatEmailDate_(booking.eventDate))}</p>
      <p><strong>Time:</strong> ${escapeEmailHtml_((booking.serviceTimes || []).join(", "))}</p>
      <p><strong>Location:</strong> ${escapeEmailHtml_(booking.location || "")}</p>
      <p><strong>Floor:</strong> ${escapeEmailHtml_(booking.floor || "")}</p>
      <p><strong>Guests:</strong> ${escapeEmailHtml_(booking.pax || "")}</p>
    </div>

    <p>
      If you think this has been done in error, please do get in touch with us as soon as possible, otherwise no further action is required.<br>
    </p>

    <p style="margin-top:24px;">
      Kind regards,<br>
      <strong>FIKA Hospitality</strong>
    </p>
  </div>
  `;
}
