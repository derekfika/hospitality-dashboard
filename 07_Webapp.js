function getDashboardBookings() {
  const sh = getDashboardSheet_();
  const data = sh.getDataRange().getValues();

  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  return rows
    .filter(row => row.some(cell => cell !== "" && cell !== null))
    .map((row, index) => {
      const obj = {};

      headers.forEach((header, i) => {
        obj[header] = normaliseForClient_(row[i]);
      });

      obj.RowNumber = index + 2;

      obj.ServiceTimesParsed = safeJsonParse_(obj.ServiceTimes, []);
      obj.ItemsParsed = safeJsonParse_(obj.ItemsJSON, []);
      obj.ParsedBooking = safeJsonParse_(obj.ParsedJSON, {});

      obj.ValidationErrorsParsed = String(obj.ValidationErrors || "")
        .split(",")
        .map(s => s.trim())
        .filter(Boolean);

      return obj;
    });
}

function normaliseForClient_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value.toISOString();
  }

  if (value === null || value === undefined) return "";

  return value;
}

function safeJsonParse_(value, fallback) {
  try {
    if (!value) return fallback;
    if (typeof value !== "string") return value;
    return JSON.parse(value);
  } catch (e) {
    return fallback;
  }
}

function doGet() {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle(getConfiguredValue_("APP_NAME", CONFIG.APP_NAME))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function updateBookingFromDashboard(rowNumber, patch) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const parsedJsonCol = map.ParsedJSON;
  if (!parsedJsonCol) throw new Error("ParsedJSON column not found.");

  const currentJson = sh.getRange(rowNumber, parsedJsonCol).getValue();
  let booking = safeJsonParse_(currentJson, null);

  if (!booking) {
    throw new Error("Could not read booking JSON for row " + rowNumber);
  }

  const oldEventDate = booking.eventDate || "";
  const oldServiceTimes = JSON.stringify(booking.serviceTimes || []);

  Object.keys(patch).forEach(key => {
    booking[key] = patch[key];
  });

  const newEventDate = booking.eventDate || "";
  const newServiceTimes = JSON.stringify(booking.serviceTimes || []);

  const dateOrTimeChanged =
    oldEventDate !== newEventDate ||
    oldServiceTimes !== newServiceTimes;

  booking.manuallyEdited = true;
  booking.lastEditedBy = Session.getActiveUser().getEmail();
  booking.lastEditedAt = new Date();
  booking.updatedAt = new Date();

  if (booking.quoteUrl) {
    booking.quoteStale = true;
  }

  if (
    dateOrTimeChanged &&
    (booking.calendarEventId || booking.calendarEventUrl)
  ) {
    booking.calendarStale = true;
  }

  booking = validateBooking_(booking);

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return {
    ok: true,
    booking
  };
}

function writeBookingObjectToExistingRow_(rowNumber, booking) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const values = {
    Status: booking.status,
    ValidationErrors: booking.validationErrors.join(", "),
    ClientCompany: booking.clientCompany,
    HostName: booking.hostName,
    HostEmail: booking.hostEmail,
    InvoiceReference: booking.invoiceReference || "",
    Pax: booking.pax,
    EventDate: booking.eventDate,
    ServiceTimes: JSON.stringify(booking.serviceTimes || []),
    ServiceType: booking.serviceType,
    Location: booking.location,
    Floor: booking.floor,
    Notes: booking.notes,
    TotalPrice: booking.totalPrice,
    MgmtFee: booking.mgmtFee,
    NetPrice: booking.netPrice,
    Vat: booking.vat,
    GrossPrice: booking.grossPrice,
    ItemsJSON: JSON.stringify(booking.items || []),
    ParsedJSON: JSON.stringify(booking),
    ManuallyEdited: booking.manuallyEdited,
    LastEditedBy: booking.lastEditedBy,
    LastEditedAt: booking.lastEditedAt,
    UpdatedAt: booking.updatedAt,
    Error: booking.error || ""
  };

  const lastCol = sh.getLastColumn();
  const currentRow = sh.getRange(rowNumber, 1, 1, lastCol).getValues()[0];

  Object.keys(values).forEach(key => {
    if (!map[key]) return;
    currentRow[map[key] - 1] = values[key];
  });

  sh.getRange(rowNumber, 1, 1, lastCol).setValues([currentRow]);

  syncBookingJsonFileIfPresent_(booking);
}

function syncBookingJsonFileIfPresent_(booking) {
  const fileId =
    booking.bookingJsonFileId ||
    extractDriveIdFromUrl_(booking.bookingJsonFileUrl || "");
  if (!fileId) return;

  try {
    const file = DriveApp.getFileById(fileId);
    updateBookingJsonFile_(file, booking);
  } catch (error) {
    console.warn(
      "Booking JSON file could not be updated for " +
      (booking.bookingId || "unknown booking") +
      ": " +
      error.message
    );
  }
}

function confirmBookingForRow(rowNumber) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");

  if (!booking.quoteUrl) {
    throw new Error("Cannot confirm booking before a quote has been generated.");
  }

  if (
    getConfiguredValue_("REQUIRE_CALENDAR_BEFORE_CONFIRMATION", true) &&
    !booking.calendarEventId
  ) {
    throw new Error("Cannot confirm booking before a calendar event has been created.");
  }
  const emailResult = sendBookingConfirmationEmail_(booking);

  booking.status = CONFIG.STATUS.CONFIRMED || "CONFIRMED";
  booking.confirmedAt = new Date();
  booking.confirmedBy = Session.getActiveUser().getEmail();
  booking.confirmationEmailSentAt = new Date();
  booking.confirmationEmailSentTo = emailResult.sentTo;
  booking.updatedAt = new Date();

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return {
    ok: true,
    bookingId: booking.bookingId,
    sentTo: emailResult.sentTo
  };
}

function setBookingStatus(rowNumber, status) {
  const allowed = Object.values(CONFIG.STATUS);
  if (!allowed.includes(status)) {
    throw new Error("Invalid status: " + status);
  }

  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");

  booking.status = status;
  booking.updatedAt = new Date();

  if (status === CONFIG.STATUS.CONFIRMED) {
    booking.confirmedAt = new Date();
  }

  if (status === CONFIG.STATUS.CANCELLED) {
    booking.cancelledAt = new Date();
  }

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return { ok: true };
}
