function generateQuoteForRow(rowNumber) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");

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
    const template = DriveApp.getFileById(CONFIG.QUOTE_TEMPLATE_DOC_ID);
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

function generateQuote(bookingId) {
  const row = BOOKINGS.find(x => x.BookingID === bookingId);
  if (!row) return;

  setBusy(bookingId);

  google.script.run
    .withSuccessHandler(res => {
      showToast("Quote generated.", "success");
      clearBusy();
      loadBookings();
      if (res.quoteUrl) window.open(res.quoteUrl, "_blank");
    })
    .withFailureHandler(err => {
      clearBusy();
      showToast("Quote failed: " + (err.message || err), "error");
    })
    .generateQuoteForRow(row.RowNumber);
}

function clearAndRefillQuoteDoc_(doc, booking) {
  const body = doc.getBody();

  const replacements = {
    "<FLOOR>": booking.floor || "",
    "<DAY>": formatDayName_(booking.eventDate),
    "<PAX>": booking.pax || "",
    "<DATE>": formatUkDate_(booking.eventDate),
    "<SERVICE>": booking.serviceType || "",
    "<DELIVERYTIME>": (booking.serviceTimes || []).join(" / "),
    "<SERVICETIME>": (booking.serviceTimes || []).join(" / "),
    "<LOCATION>": booking.location || "",
    "<HOST>": booking.hostName || booking.hostEmail || "",
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
      const detail = body.insertParagraph(insertIndex++, detailParts.join(" — "));
      detail.setFontSize(8);
      detail.setItalic(true);
      detail.setSpacingAfter(4);
    }
  });
}

function getQuoteFolderForBooking_(booking) {
  const root = getOrCreateDriveFolder_(CONFIG.QUOTE_ROOT_FOLDER_NAME || "Hospitality");

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
  return "£" + Number(n || 0).toFixed(2);
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
    to: CONFIG.PRINTER_EMAIL,
    subject: "Angel Court Hospitality Quote",
    body: "Auto-printed quote from Angel Court Hospitality Dashboard.",
    attachments: [pdfBlob]
  });

  booking.quotePrintedAt = new Date();
  booking.status = CONFIG.STATUS.PRINTED;
  booking.updatedAt = new Date();

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return {
    ok: true
  };
}