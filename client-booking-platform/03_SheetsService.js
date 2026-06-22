const BOOKING_HEADERS = [
  "Booking ID", "Submitted At", "Status", "Source", "Site", "Event Date", "Start Time", "End Time",
  "Guest Count", "Client Name", "Client Email", "Client Phone", "Company Name", "Floor Level",
  "Room / Area", "Net Total", "Dietary Summary", "Warnings Summary", "Booking JSON", "Processed",
  "Quote Created", "Calendar Created", "Kitchen Printed", "Internal Notes"
];

const LINE_ITEM_HEADERS = [
  "Booking ID", "Category", "Item ID", "Item Name", "Serving Info", "Unit Price", "Quantity",
  "Line Total", "Time Required", "Choices", "Comments", "Notice Required Days", "Minimum Order"
];

const SETTINGS_HEADERS = ["Key", "Value", "Section", "Label", "Notes"];

const PLATFORM_SETTINGS_ROWS = [
  ["FIKA_LOGO_URL", "", "Branding", "FIKA logo image URL", "Use a direct HTTPS image URL. Leave blank to show the fallback wordmark."],
  ["FIKA_LOGO_ALT", "FIKA", "Branding", "FIKA logo alternative text", "Used for accessibility when the image is present."],
  ["FIKA_FALLBACK_TEXT", "Fika", "Branding", "FIKA fallback wordmark", "Shown when the image URL is blank or cannot load."],
  ["SITE_LOGO_URL", "", "Branding", "Angel Court logo image URL", "Use a direct HTTPS image URL. Leave blank to show the fallback wordmark."],
  ["SITE_LOGO_ALT", "Angel Court", "Branding", "Site logo alternative text", "Used for accessibility when the image is present."],
  ["SITE_FALLBACK_TEXT", "Angel Court", "Branding", "Site fallback wordmark", "Shown when the image URL is blank or cannot load."],
  ["CLIENT_FACING_NAME", SITE_CONFIG.clientFacingName, "Copy", "Browser and platform name", "The client-facing name of the booking platform."],
  ["BRAND_EYEBROW", SITE_CONFIG.branding.eyebrow, "Copy", "Hero eyebrow", "Small heading above the main title."],
  ["HERO_TITLE", SITE_CONFIG.branding.heroTitle, "Copy", "Hero title", "Main booking page headline."],
  ["HERO_BODY", SITE_CONFIG.branding.heroBody, "Copy", "Hero body copy", "Short introduction below the headline."],
  ["COLOUR_ACCENT", SITE_CONFIG.branding.accent, "Colours", "Primary brand colour", "Hex colour such as #3d21bf."],
  ["COLOUR_INK", SITE_CONFIG.branding.ink, "Colours", "Text colour", "Hex colour such as #221874."],
  ["COLOUR_PAPER", SITE_CONFIG.branding.paper, "Colours", "Page background colour", "Hex colour such as #f4f4f2."]
];

const DASHBOARD_REQUIRED_HEADERS = [
  "BookingID", "Status", "ValidationErrors", "ClientCompany", "HostName", "HostEmail", "Pax",
  "EventDate", "ServiceTimes", "ServiceType", "Location", "Floor", "Notes", "TotalPrice",
  "MgmtFee", "NetPrice", "Vat", "GrossPrice", "ItemsJSON", "ParsedJSON"
];

function setupBookingPlatformSheets() {
  const spreadsheet = getBookingSpreadsheet_();
  const result = { ok: true, sheets: [] };

  if (SITE_CONFIG.integration.keepClientRequestLog) {
    const requests = getOrCreateSheet_(spreadsheet, SITE_CONFIG.sheets.bookingRequests, BOOKING_HEADERS);
    requests.setFrozenRows(1);
    result.sheets.push(requests.getName());
  }

  const lines = getOrCreateSheet_(spreadsheet, SITE_CONFIG.sheets.bookingLineItems, LINE_ITEM_HEADERS);
  lines.setFrozenRows(1);
  result.sheets.push(lines.getName());

  const settings = setupPlatformSettingsSheet_(spreadsheet);
  result.sheets.push(settings.getName());

  const dashboard = getDashboardDataSheet_(spreadsheet);
  const map = getSheetHeaderMap_(dashboard);
  assertHeaders_(map, DASHBOARD_REQUIRED_HEADERS, dashboard.getName());
  result.sheets.push(dashboard.getName());
  return result;
}

function setupPlatformSettingsSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SITE_CONFIG.sheets.settings);
  if (!sheet) sheet = spreadsheet.insertSheet(SITE_CONFIG.sheets.settings);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setValues([SETTINGS_HEADERS]);
  }
  const existingKeys = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().map(function(row) { return row[0]; })
    : [];
  const missingRows = PLATFORM_SETTINGS_ROWS.filter(function(row) { return existingKeys.indexOf(row[0]) === -1; });
  if (missingRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missingRows.length, SETTINGS_HEADERS.length).setValues(missingRows);
  }
  sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length)
    .setFontWeight("bold").setBackground("#3d21bf").setFontColor("#ffffff");
  if (sheet.getLastRow() > 1) sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).setBackground("#e9fbf5");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, SETTINGS_HEADERS.length);
  sheet.setColumnWidth(2, 340);
  sheet.setColumnWidth(5, 420);
  return sheet;
}

function writeBookingRequest_(booking) {
  if (SITE_CONFIG.integration.mode !== "direct_dashboard") {
    throw new Error("Unsupported booking integration mode: " + SITE_CONFIG.integration.mode);
  }

  const spreadsheet = getBookingSpreadsheet_();
  const dashboardSheet = getDashboardDataSheet_(spreadsheet);
  const lineSheet = getOrCreateSheet_(spreadsheet, SITE_CONFIG.sheets.bookingLineItems, LINE_ITEM_HEADERS);
  const dashboardBooking = adaptClientBookingForDashboard_(booking);
  const dashboardRow = appendDashboardBooking_(dashboardSheet, dashboardBooking);
  const firstLineItemRow = lineSheet.getLastRow() + 1;
  let lineItemRowsWritten = 0;

  try {
    appendClientLineItems_(lineSheet, booking);
    lineItemRowsWritten = booking.order.items.length;
    if (SITE_CONFIG.integration.keepClientRequestLog) {
      appendClientRequestLog_(getOrCreateSheet_(spreadsheet, SITE_CONFIG.sheets.bookingRequests, BOOKING_HEADERS), booking);
    }
  } catch (error) {
    if (lineItemRowsWritten) lineSheet.deleteRows(firstLineItemRow, lineItemRowsWritten);
    dashboardSheet.deleteRow(dashboardRow);
    throw error;
  }

  return {
    dashboardRow: dashboardRow,
    dashboardStatus: dashboardBooking.status
  };
}

function adaptClientBookingForDashboard_(booking) {
  const subtotal = roundMoney_(booking.order.netTotal);
  const managementFee = roundMoney_(subtotal * Number(SITE_CONFIG.integration.managementFeeRate || 0));
  const netPrice = roundMoney_(subtotal + managementFee);
  const vat = roundMoney_(netPrice * Number(SITE_CONFIG.integration.vatRate || 0));
  const grossPrice = roundMoney_(netPrice + vat);
  const serviceType = eventTypeLabel_(booking.order.eventType);
  const serviceTimes = unique_(booking.order.items.map(function(item) {
    return item.timeRequired || booking.event.startTime;
  }).filter(Boolean));
  if (!serviceTimes.length && booking.event.startTime) serviceTimes.push(booking.event.startTime);

  const dashboardItems = booking.order.items.map(function(item) {
    const choiceText = item.choices.filter(function(choice) { return choice.value; })
      .map(function(choice) { return choice.label + ": " + choice.value; }).join("; ");
    return {
      section: item.category,
      name: item.itemName,
      detail: choiceText,
      info: item.servingInfo,
      qty: item.quantity,
      time: item.timeRequired || booking.event.startTime,
      comment: item.comments,
      unitPrice: item.unitPrice,
      lineTotal: item.lineTotal,
      itemId: item.itemId
    };
  });

  const notes = buildDashboardNotes_(booking);
  const validationErrors = [];
  if (!booking.client.companyName) validationErrors.push("Missing company");
  if (!booking.event.eventDate) validationErrors.push("Missing event date");
  if (!serviceTimes.length) validationErrors.push("Missing service time");
  if (!booking.event.guestCount) validationErrors.push("Missing pax");
  if (!dashboardItems.length) validationErrors.push("Missing line items");
  if (subtotal <= 0) validationErrors.push("Missing total price");
  return {
    bookingId: booking.bookingId,
    status: validationErrors.length ? "NEEDS_REVIEW" : "READY",
    validationErrors: validationErrors,
    emailReceived: booking.submittedAt,
    messageId: "CLIENT:" + booking.bookingId,
    threadId: "",
    attachmentName: "client-booking-" + booking.bookingId + ".json",
    sourceEmailFrom: booking.client.email,
    sourceEmailSubject: "Client booking request " + booking.bookingId,
    source: booking.source,
    sourceType: "CLIENT_PLATFORM",
    clientCompany: booking.client.companyName,
    hostName: booking.client.name,
    hostEmail: booking.client.email,
    hostPhone: booking.client.phone,
    pax: booking.event.guestCount,
    eventDate: booking.event.eventDate,
    serviceTimes: serviceTimes,
    serviceType: serviceType,
    location: SITE_CONFIG.address,
    floor: booking.event.floorLevel,
    notes: notes,
    totalPrice: subtotal,
    mgmtFee: managementFee,
    netPrice: netPrice,
    vat: vat,
    grossPrice: grossPrice,
    items: dashboardItems,
    clientBooking: booking,
    quoteUrl: "",
    quoteCreatedAt: "",
    quotePrintedAt: "",
    calendarEventId: "",
    calendarEventUrl: "",
    calendarCreatedAt: "",
    manuallyEdited: false,
    lastEditedBy: "",
    lastEditedAt: "",
    createdAt: booking.submittedAt,
    updatedAt: booking.submittedAt,
    error: "",
    quoteStale: false,
    calendarStale: false
  };
}

function appendDashboardBooking_(sheet, booking) {
  const map = getSheetHeaderMap_(sheet);
  assertHeaders_(map, DASHBOARD_REQUIRED_HEADERS, sheet.getName());
  assertDashboardBookingIdIsUnique_(sheet, map, booking.bookingId);

  const values = {
    BookingID: booking.bookingId,
    Status: booking.status,
    ValidationErrors: booking.validationErrors.join(", "),
    EmailReceived: new Date(booking.emailReceived),
    MessageId: booking.messageId,
    ThreadId: booking.threadId,
    AttachmentName: booking.attachmentName,
    SourceEmailFrom: booking.sourceEmailFrom,
    SourceEmailSubject: booking.sourceEmailSubject,
    ClientCompany: booking.clientCompany,
    HostName: booking.hostName,
    HostEmail: booking.hostEmail,
    Pax: booking.pax,
    EventDate: booking.eventDate,
    ServiceTimes: JSON.stringify(booking.serviceTimes),
    ServiceType: booking.serviceType,
    Location: booking.location,
    Floor: booking.floor,
    Notes: booking.notes,
    TotalPrice: booking.totalPrice,
    MgmtFee: booking.mgmtFee,
    NetPrice: booking.netPrice,
    Vat: booking.vat,
    GrossPrice: booking.grossPrice,
    ItemsJSON: JSON.stringify(booking.items),
    ParsedJSON: JSON.stringify(booking),
    QuoteURL: "",
    QuoteCreatedAt: "",
    QuotePrintedAt: "",
    CalendarEventId: "",
    CalendarEventURL: "",
    CalendarCreatedAt: "",
    ManuallyEdited: false,
    LastEditedBy: "",
    LastEditedAt: "",
    CreatedAt: new Date(booking.createdAt),
    UpdatedAt: new Date(booking.updatedAt),
    Error: ""
  };

  const rowNumber = sheet.getLastRow() + 1;
  const row = new Array(sheet.getLastColumn()).fill("");
  Object.keys(values).forEach(function(header) {
    if (map[header]) row[map[header] - 1] = values[header];
  });
  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  return rowNumber;
}

function appendClientLineItems_(sheet, booking) {
  if (!booking.order.items.length) return;
  const rows = booking.order.items.map(function(item) {
    return [
      booking.bookingId, item.category, item.itemId, item.itemName, item.servingInfo,
      item.unitPrice, item.quantity, item.lineTotal, item.timeRequired, JSON.stringify(item.choices),
      item.comments, item.noticeRequiredDays, item.minimumOrder
    ];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, LINE_ITEM_HEADERS.length).setValues(rows);
}

function appendClientRequestLog_(sheet, booking) {
  sheet.appendRow([
    booking.bookingId, new Date(booking.submittedAt), booking.status, booking.source, booking.site,
    booking.event.eventDate, booking.event.startTime, booking.event.endTime, booking.event.guestCount,
    booking.client.name, booking.client.email, booking.client.phone, booking.client.companyName,
    booking.event.floorLevel, booking.event.roomOrArea || booking.event.deliveryPoint, booking.order.netTotal,
    dietarySummary_(booking.dietaries), warningSummary_(booking.warnings), JSON.stringify(booking),
    false, false, false, false, ""
  ]);
}

function getBookingSpreadsheet_() {
  const id = String(SITE_CONFIG.integration.spreadsheetId || "").trim();
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getDashboardDataSheet_(spreadsheet) {
  const name = SITE_CONFIG.integration.dashboardSheetName;
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error("Dashboard sheet '" + name + "' was not found. Check SITE_CONFIG.integration.");
  return sheet;
}

function getOrCreateSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#3d21bf").setFontColor("#ffffff");
  }
  return sheet;
}

function getSheetHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(header, index) { map[String(header).trim()] = index + 1; });
  return map;
}

function assertHeaders_(map, required, sheetName) {
  const missing = required.filter(function(header) { return !map[header]; });
  if (missing.length) throw new Error("Sheet '" + sheetName + "' is missing headers: " + missing.join(", "));
}

function assertDashboardBookingIdIsUnique_(sheet, map, bookingId) {
  if (sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, map.BookingID, sheet.getLastRow() - 1, 1).getDisplayValues();
  if (values.some(function(row) { return row[0] === bookingId; })) {
    throw new Error("Booking " + bookingId + " already exists in the dashboard.");
  }
}

function buildDashboardNotes_(booking) {
  const dietary = dietarySummary_(booking.dietaries);
  return [
    "Submitted through Client Booking Platform.",
    booking.specialInstructions ? "Special instructions: " + booking.specialInstructions : "",
    booking.event.roomOrArea ? "Room / area: " + booking.event.roomOrArea : "",
    booking.event.deliveryPoint ? "Delivery point: " + booking.event.deliveryPoint : "",
    dietary !== "None declared" ? "Dietaries: " + dietary : "Dietaries: None declared",
    booking.dietaries.allergyDetails ? "ALLERGIES: " + booking.dietaries.allergyDetails : "",
    warningSummary_(booking.warnings) ? "Warnings: " + warningSummary_(booking.warnings) : "",
    booking.event.onsiteContactName ? "Onsite contact: " + booking.event.onsiteContactName + " " + booking.event.onsiteContactPhone : ""
  ].filter(Boolean).join("\n");
}

function eventTypeLabel_(id) {
  const type = EVENT_TYPES.find(function(candidate) { return candidate.id === id; });
  return type ? type.label : id;
}

function unique_(values) {
  return values.filter(function(value, index) { return values.indexOf(value) === index; });
}

function dietarySummary_(dietaries) {
  if (!dietaries.hasDietaries && !dietaries.allergyDetails) return "None declared";
  return [
    "Vegetarian " + dietaries.vegetarian, "Vegan " + dietaries.vegan,
    "Gluten-free " + dietaries.glutenFree, "Coeliac " + dietaries.coeliac,
    "Dairy-free " + dietaries.dairyFree, "Halal " + dietaries.halal,
    "Other " + dietaries.otherCount
  ].join(" · ");
}

function warningSummary_(warnings) {
  const output = [];
  if (warnings.inside72Hours) output.push("Inside 72 hours");
  if (warnings.inside10WorkingDays) output.push("Inside 10 working days");
  if (warnings.insideDietaryDeadline) output.push("Inside dietary deadline");
  return output.concat(warnings.minimumOrderIssues, warnings.dietaryCountIssues, warnings.itemNoticeIssues).join(" · ");
}
