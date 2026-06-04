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

function cancelBookingForRow(rowNumber) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");

  if (booking.calendarEventId) {
    try {
      Calendar.Events.remove(
        CONFIG.CALENDAR_ID || "primary",
        booking.calendarEventId,
        {
          sendUpdates: "all"
        }
      );
    } catch (e) {
      Logger.log("Calendar deletion failed: " + e);
    }
  }

  booking.status = CONFIG.STATUS.CANCELLED;
  booking.cancelledAt = new Date();
  booking.calendarEventId = "";
  booking.calendarEventUrl = "";
  booking.calendarStale = false;
  booking.updatedAt = new Date();

  writeBookingObjectToExistingRow_(rowNumber, booking);

  return { ok: true };
}

function setBusy(bookingId) {
  BUSY_BOOKING_ID = bookingId;
  renderBookings();
}

function clearBusy() {
  BUSY_BOOKING_ID = "";
  renderBookings();
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

function testGetDashboardBookings() {
  const bookings = getDashboardBookings();
  Logger.log(JSON.stringify(bookings, null, 2));
}

function doGet() {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Angel Court Hospitality Dashboard")
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

  Object.keys(values).forEach(key => {
    if (!map[key]) return;
    sh.getRange(rowNumber, map[key]).setValue(values[key]);
  });
}

function confirmBookingForRow(rowNumber) {
  const sh = getDashboardSheet_();
  const map = getHeaderMap_();

  const json = sh.getRange(rowNumber, map.ParsedJSON).getValue();
  let booking = safeJsonParse_(json, null);

  if (!booking) throw new Error("Could not read booking data.");

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

function setSort(key) {
  if (SORT_KEY === key) {
    SORT_DIR = SORT_DIR === "asc" ? "desc" : "asc";
  } else {
    SORT_KEY = key;
    SORT_DIR = "asc";
  }

  renderBookings();
}

function getSortValue(b, key) {
  switch (key) {
    case "date": return getDisplayDate(b) || "";
    case "time": return getDisplayTime(b) || "";
    case "company": return String(b.ClientCompany || "").toLowerCase();
    case "pax": return Number(b.Pax || 0);
    case "service": return String(b.ServiceType || "").toLowerCase();
    case "floor": return String(b.Floor || "").toLowerCase();
    case "total": return Number(b.TotalPrice || 0);
    case "status": return String(b.Status || "").toLowerCase();
    default: return "";
  }
}

function isBeforeThisWeek_(booking) {
  if (!booking.eventDate) return false;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const day = today.getDay(); // Sun 0, Mon 1
  const diffToMonday = day === 0 ? 6 : day - 1;

  const monday = new Date(today);
  monday.setDate(today.getDate() - diffToMonday);

  const parts = String(booking.eventDate).split("-");
  if (parts.length !== 3) return false;

  const bookingDate = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );
  bookingDate.setHours(0, 0, 0, 0);

  return bookingDate < monday;
}