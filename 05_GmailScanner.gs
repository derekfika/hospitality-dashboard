function scanInboxForDashboardBookings() {
  const query = buildInboxScanQuery_();
  console.log("Inbox scan query: " + query);

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
          if (shouldArchiveBooking_(booking)) {
            booking.status = CONFIG.STATUS.ARCHIVED || "ARCHIVED";
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
  setSetting_("LAST_INBOX_SCAN_AT", new Date().toISOString());
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

function buildInboxScanQuery_() {
  const processedLabel =
    getConfiguredValue_("PROCESSED_LABEL_NAME", "AC_HOSPITALITY_PROCESSED");

  let query =
    `in:anywhere -in:trash -in:spam filename:xlsx -label:${processedLabel}`;

  const earliest =
    normaliseSettingsDate_(getConfiguredValue_("EARLIEST_SCAN_DATE", ""));

  const lastScan =
    normaliseSettingsDate_(getConfiguredValue_("LAST_INBOX_SCAN_AT", ""));

  const scanAfter = getLaterDate_(earliest, lastScan);

  if (scanAfter) {
    query += " after:" + Utilities.formatDate(
      scanAfter,
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    );
  } else {
    query += " newer_than:90d";
  }

  return query;
}

function normaliseSettingsDate_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value;
  }

  const text = String(value).trim();
  if (!text) return null;

  let m = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  m = text.match(/^(\d{4})\/(\d{2})\/(\d{2})/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  return null;
}

function getLaterDate_(a, b) {
  if (a && b) return a > b ? a : b;
  return a || b || null;
}

function applyProcessedLabel_(thread) {
  const labelName = getConfiguredValue_("PROCESSED_LABEL_NAME", "AC_HOSPITALITY_PROCESSED");

  let label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    label = GmailApp.createLabel(labelName);
  }

  thread.addLabel(label);
}

function shouldArchiveBooking_(booking) {
  if (!booking.eventDate) return false;

  const archiveAfterDays = getConfiguredNumber_("ARCHIVE_AFTER_DAYS", 0);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return isBookingOlderThanArchiveThreshold_(booking.eventDate, archiveAfterDays, today);
}

function isBookingOlderThanArchiveThreshold_(eventDate, archiveAfterDays, today) {
  if (!eventDate) return false;

  const threshold = new Date(today);
  threshold.setDate(threshold.getDate() - Number(archiveAfterDays || 0));

  const parts = String(eventDate).split("-");
  if (parts.length !== 3) return false;

  const bookingDate = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );
  bookingDate.setHours(0, 0, 0, 0);

  return bookingDate < threshold;
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

    booking = normaliseBookingTimes_(booking);
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

function prepareInboxScan() {
  const query = buildInboxScanQuery_();
  const threads = GmailApp.search(query, 0, 500);

  const scanFrom =
    getConfiguredValue_("LAST_INBOX_SCAN_AT", "") ||
    getConfiguredValue_("EARLIEST_SCAN_DATE", "") ||
    "recent inbox";

  return {
    query,
    totalFound: threads.length,
    scanFromLabel: String(scanFrom)
  };
}

function scanInboxChunk(offset, limit) {
  offset = Number(offset || 0);
  limit = Number(limit || 5);

  const query = buildInboxScanQuery_();
  const threads = GmailApp.search(query, offset, limit);

  let logged = 0;
  let skipped = 0;
  let errors = 0;

  threads.forEach(thread => {
    try {
      const result = processInboxThread_(thread);

      if (result && result.logged) logged += result.logged;
      else skipped++;
    } catch (e) {
      console.log("Scan thread error: " + e);
      errors++;
    }
  });

  const nextOffset = offset + threads.length;
  const done = threads.length < limit;

  if (done) {
    setSetting_("LAST_INBOX_SCAN_AT", new Date().toISOString());
  }

  return {
    logged,
    skipped,
    errors,
    nextOffset,
    done
  };
}

function processInboxThread_(thread) {
  let logged = 0;
  let skipped = 0;
  let errors = 0;
  let scanned = 0;

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
        const booking = buildBookingFromMessageIdAndAttachment_(
          msg.getId(),
          att.getName()
        );

        if (shouldArchiveBooking_(booking)) {
          booking.status = CONFIG.STATUS.ARCHIVED || "ARCHIVED";
        }

        writeBookingToSheet_(booking);
        logged++;

      } catch (e) {
        console.log(e);
        errors++;
      }
    }
  }

  return {
    scanned,
    logged,
    skipped,
    errors
  };
}
