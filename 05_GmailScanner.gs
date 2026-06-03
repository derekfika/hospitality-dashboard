function scanInboxForDashboardBookings() {
  const lastScan = getSetting_("LAST_INBOX_SCAN_AT", "");

  let query =
    'in:anywhere -in:trash -in:spam filename:xlsx -label:AC_HOSPITALITY_PROCESSED';

  if (lastScan) {
    const afterDate = Utilities.formatDate(
      new Date(lastScan),
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    );

    query += ` after:${afterDate}`;
  } else {
    query += " newer_than:90d";
  }

  const threads = GmailApp.search(query, 0, 50);

  let scanned = 0;
  let logged = 0;
  let skipped = 0;
  let errors = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const msg of messages) {
      const atts = msg.getAttachments({ includeInlineImages: false });

      for (const att of atts) {
        if (!isXlsx_(att)) continue;

        scanned++;

        const key = makeDashboardKey_(msg.getId(), att.getName());

        if (isDashboardKeyLogged_(key)) {
          skipped++;
          continue;
        }

        try {
          const booking = buildBookingFromMessageIdAndAttachment_(msg.getId(), att.getName());
          if (isPastBooking_(booking)) {
            booking.status = CONFIG.STATUS.ARCHIVED || "ARCHIVED";
            applyProcessedLabel_(thread);
            skipped++;
            continue;
          }
          writeBookingToSheet_(booking);
          logged++;
        } catch (e) {
          errors++;
          Logger.log("Booking scan error: " + e);
        }
      }
    }
  }

  Logger.log(
    `Dashboard scan complete. Scanned=${scanned}, Logged=${logged}, Skipped=${skipped}, Errors=${errors}`
  );

  return {
    scanned,
    logged,
    skipped,
    errors
  };
}

function applyProcessedLabel_(thread) {
  const labelName = "AC_HOSPITALITY_PROCESSED";

  let label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    label = GmailApp.createLabel(labelName);
  }

  thread.addLabel(label);
}

function isPastBooking_(booking) {
  if (!booking.eventDate) return false;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const parts = String(booking.eventDate).split("-");
  if (parts.length !== 3) return false;

  const d = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );

  d.setHours(0, 0, 0, 0);

  return d < today;
}

function makeDashboardKey_(messageId, attachmentName) {
  return `${messageId}|${attachmentName || ""}`;
}

function isDashboardKeyLogged_(key) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const messageCol = map.MessageId;
  const attachmentCol = map.AttachmentName;

  if (!messageCol || !attachmentCol) {
    throw new Error("MessageId or AttachmentName column missing.");
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const values = sh
    .getRange(2, 1, lastRow - 1, sh.getLastColumn())
    .getValues();

  for (const row of values) {
    const existingKey =
      `${row[messageCol - 1]}|${row[attachmentCol - 1] || ""}`;

    if (existingKey === key) return true;
  }

  return false;
}

function buildBookingFromMessageIdAndAttachment_(messageId, attachmentName) {
  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  const atts = msg.getAttachments({ includeInlineImages: false });

  let xlsx = null;

  for (const att of atts) {
    if (!isXlsx_(att)) continue;

    if (att.getName() === attachmentName) {
      xlsx = att;
      break;
    }

    if (!xlsx) xlsx = att;
  }

  if (!xlsx) throw new Error("No XLSX attachment found.");

  let tempSheetId = null;

  try {
    tempSheetId = convertXlsxToGoogleSheet_(xlsx);

    const ss = SpreadsheetApp.openById(tempSheetId);
    const sheet = ss.getSheets()[0];

    let booking = parseAngelCourtBookingSheet_(sheet);

    if (!booking.serviceTimes || booking.serviceTimes.length === 0) {
      const fallbackTime =
        extractTimeFromText_(msg.getSubject()) ||
        extractTimeFromText_(xlsx.getName()) ||
        extractTimeFromText_(sheet.getName());

      if (fallbackTime) {
        booking.serviceTimes = [fallbackTime];

        booking.items = booking.items.map(item => {
          if (!item.time) item.time = fallbackTime;
          return item;
        });
      }
    }

    booking.bookingId = generateBookingId_();
    booking.messageId = msg.getId();
    booking.threadId = thread.getId();
    booking.attachmentName = xlsx.getName();
    booking.emailReceived = msg.getDate();
    booking.sourceEmailFrom = msg.getFrom();
    booking.sourceEmailSubject = msg.getSubject();

    booking.hostEmail = booking.hostEmail || extractEmailAddress_(msg.getFrom());

    if (!booking.hostName) {
      booking.hostName = extractEmailName_(msg.getFrom());
    }

    booking = validateBooking_(booking);

    return booking;

  } finally {
    if (tempSheetId) {
      try {
        DriveApp.getFileById(tempSheetId).setTrashed(true);
      } catch (e) { }
    }
  }
}