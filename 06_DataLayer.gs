// =========================
// DATA LAYER
// =========================

function getDashboardSheet_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sh) {
    throw new Error(
      `Sheet '${CONFIG.SHEET_NAME}' not found`
    );
  }

  return sh;
}


// =========================
// HEADER MAP
// =========================

function getHeaderMap_() {

  const sh = getDashboardSheet_();

  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0];

  const map = {};

  headers.forEach((header, i) => {
    map[String(header).trim()] = i + 1;
  });

  return map;
}


// =========================
// WRITE BOOKING
// =========================

function writeBookingToSheet_(booking) {

  const sh = getDashboardSheet_();

  const map = getHeaderMap_();

  const row = sh.getLastRow() + 1;

  const values = {};

  values.BookingID = booking.bookingId;

  values.Status = booking.status;

  values.ValidationErrors =
    booking.validationErrors.join(", ");

  values.EmailReceived =
    booking.emailReceived;

  values.MessageId =
    booking.messageId;

  values.ThreadId =
    booking.threadId;

  values.AttachmentName =
    booking.attachmentName;

  values.SourceEmailFrom =
    booking.sourceEmailFrom;

  values.SourceEmailSubject =
    booking.sourceEmailSubject;

  values.ClientCompany =
    booking.clientCompany;

  values.HostName =
    booking.hostName;

  values.HostEmail =
    booking.hostEmail;

  values.Pax =
    booking.pax;

  values.EventDate =
    booking.eventDate;

  values.ServiceTimes =
    JSON.stringify(booking.serviceTimes);

  values.ServiceType =
    booking.serviceType;

  values.Location =
    booking.location;

  values.Floor =
    booking.floor;

  values.Notes =
    booking.notes;

  values.TotalPrice =
    booking.totalPrice;

  values.MgmtFee =
    booking.mgmtFee;

  values.NetPrice =
    booking.netPrice;

  values.Vat =
    booking.vat;

  values.GrossPrice =
    booking.grossPrice;

  values.ItemsJSON =
    JSON.stringify(booking.items);

  values.ParsedJSON =
    JSON.stringify(booking);

  values.QuoteURL =
    booking.quoteUrl;

  values.QuoteCreatedAt =
    booking.quoteCreatedAt;

  values.QuotePrintedAt =
    booking.quotePrintedAt;

  values.CalendarEventId =
    booking.calendarEventId;

  values.CalendarEventURL =
    booking.calendarEventUrl;

  values.CalendarCreatedAt =
    booking.calendarCreatedAt;

  values.ManuallyEdited =
    booking.manuallyEdited;

  values.LastEditedBy =
    booking.lastEditedBy;

  values.LastEditedAt =
    booking.lastEditedAt;

  values.CreatedAt =
    booking.createdAt;

  values.UpdatedAt =
    booking.updatedAt;

  values.Error =
    booking.error;


  // Write values into correct columns

  Object.keys(values).forEach(key => {

    if (!map[key]) return;

    sh.getRange(row, map[key])
      .setValue(values[key]);

  });

}

function testWriteBooking() {

  let booking = createEmptyBooking_();

  booking.bookingId =
    generateBookingId_();

  booking.clientCompany =
    "TEST COMPANY";

  booking.pax = 12;

  booking.eventDate =
    "2026-05-29";

  booking.serviceTimes =
    ["11:30"];

  booking.location =
    "Floor 6";

  booking.serviceType =
    "Lunch";

  booking.totalPrice = 120;

  booking.items = [
    {
      name: "Working Lunch",
      qty: 12
    }
  ];

  booking = validateBooking_(booking);

  writeBookingToSheet_(booking);

}