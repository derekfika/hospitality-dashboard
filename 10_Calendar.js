function createCalendarEventForRow(rowNumber) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");
  if (booking.calendarEventId || booking.calendarEventUrl) {
  throw new Error("Calendar event already exists for this booking.");
  }
  if (!booking.quoteUrl) throw new Error("Generate quote before creating calendar event.");
  if (!booking.messageId) throw new Error("Missing source Gmail message ID.");
  if (!booking.eventDate) throw new Error("Missing event date.");
  if (!booking.serviceTimes || booking.serviceTimes.length === 0) {
    throw new Error("Missing service time.");
  }

  const quoteFileId = extractDriveIdFromUrl_(booking.quoteUrl);
  if (!quoteFileId) throw new Error("Could not read quote file ID.");

  const quoteFile = DriveApp.getFileById(quoteFileId);

  const sourceXlsxFile = saveOriginalBookingXlsxToDrive_(booking);
  const attendees = getCalendarAttendeesFromSettings_();
  const calendarId = getConfiguredValue_("CALENDAR_ID", CONFIG.CALENDAR_ID || "primary");
  const eventDuration = getConfiguredNumber_("CALENDAR_EVENT_DURATION_MINUTES", CONFIG.CALENDAR_EVENT_DURATION_MINUTES || 60);
  const eventColorId = getConfiguredValue_("CALENDAR_EVENT_COLOR_ID", CONFIG.CALENDAR_EVENT_COLOR_ID || "9");

  const start = buildCalendarStart_(booking.eventDate, booking.serviceTimes[0]);
  const end = new Date(start.getTime() + eventDuration * 60 * 1000);

  const title = makeCalendarTitle_(booking);

  const eventResource = {
    summary: title,
    location: `${booking.location || "One Angel Court"} ${booking.floor || ""}`.trim(),
    description: "",
    start: {
      dateTime: start.toISOString(),
      timeZone: Session.getScriptTimeZone()
    },
    end: {
      dateTime: end.toISOString(),
      timeZone: Session.getScriptTimeZone()
    },
    colorId: eventColorId,
    attendees: attendees,
    attachments: [
      {
        fileUrl: quoteFile.getUrl(),
        title: quoteFile.getName(),
        mimeType: quoteFile.getMimeType()
      },
      {
        fileUrl: sourceXlsxFile.getUrl(),
        title: sourceXlsxFile.getName(),
        mimeType: sourceXlsxFile.getMimeType()
      }
    ]
  };

  const created = Calendar.Events.insert(
    eventResource,
    calendarId,
    {
      supportsAttachments: true,
      sendUpdates: "all"
    }
  );

  booking.calendarEventId = created.id || "";
  booking.calendarEventUrl = created.htmlLink || "";
  booking.calendarCreatedAt = new Date();
  booking.status = CONFIG.STATUS.CPU_CREATED;
  booking.updatedAt = new Date();
  booking.calendarStale = false;
  writeBookingObjectToExistingRow_(rowNumber, booking);

  return {
    ok: true,
    eventUrl: created.htmlLink || ""
  };
}

function saveOriginalBookingXlsxToDrive_(booking) {
  const msg = GmailApp.getMessageById(booking.messageId);
  const atts = msg.getAttachments({ includeInlineImages: false });

  let chosen = null;

  for (const att of atts) {
    if (!isXlsx_(att)) continue;

    if (booking.attachmentName && att.getName() === booking.attachmentName) {
      chosen = att;
      break;
    }

    if (!chosen) chosen = att;
  }

  if (!chosen) throw new Error("Could not find original XLSX attachment.");

  const folder = getQuoteFolderForBooking_(booking);

  const fileName = `Original Booking Form - ${booking.clientCompany || "Unknown"} - ${booking.eventDate || ""}.xlsx`;

  return folder.createFile(chosen.copyBlob()).setName(fileName);
}

function getCalendarAttendeesFromSettings_() {
  return String(
    getConfiguredValue_(
      "CALENDAR_ATTENDEES",
      (CONFIG.CALENDAR_ATTENDEES || []).join(", ")
    )
  )
    .split(",")
    .map(email => email.trim())
    .filter(Boolean)
    .map(email => ({ email }));
}

function makeCalendarTitle_(booking) {
  const locationShortCode = getConfiguredValue_("LOCATION_SHORT_CODE", CONFIG.LOCATION_SHORT_CODE || "OAC");
  const company = booking.clientCompany || "Unknown Company";
  const service = booking.serviceType || "Hospitality";
  const pax = booking.pax || "";

  return `${locationShortCode}_${company}_${service} x ${pax}`;
}

function buildCalendarStart_(isoDate, timeText) {
  const parts = String(isoDate).split("-");
  if (parts.length !== 3) throw new Error("Invalid event date.");

  const timeParts = String(timeText || "00:00").split(":");

  const y = Number(parts[0]);
  const m = Number(parts[1]) - 1;
  const d = Number(parts[2]);

  const hh = Number(timeParts[0] || 0);
  const mm = Number(timeParts[1] || 0);

  const dt = new Date(y, m, d, hh, mm, 0, 0);

  if (isNaN(dt.getTime())) throw new Error("Invalid calendar start time.");

  return dt;
}

function resetCalendarForRow(rowNumber) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");

  booking.calendarEventId = "";
  booking.calendarEventUrl = "";
  booking.calendarCreatedAt = "";
  booking.calendarStale = false;

  if (booking.quoteUrl) {
    booking.status = CONFIG.STATUS.QUOTE_GENERATED;
  } else {
    booking.status = CONFIG.STATUS.READY;
  }

  booking.updatedAt = new Date();

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return { ok: true };
}
