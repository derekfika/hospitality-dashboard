function createEmptyBooking_() {

  return {

    bookingId: "",

    status: CONFIG.STATUS.NEW,

    validationErrors: [],

    emailReceived: "",
    messageId: "",
    threadId: "",
    attachmentName: "",

    sourceEmailFrom: "",
    sourceEmailSubject: "",

    clientCompany: "",
    hostName: "",
    hostEmail: "",

    pax: "",

    eventDate: "",
    serviceTimes: [],

    serviceType: "",

    location: "",
    floor: "",

    notes: "",

    totalPrice: 0,
    mgmtFee: 0,
    netPrice: 0,
    vat: 0,
    grossPrice: 0,

    items: [],

    quoteUrl: "",

    quoteCreatedAt: "",
    quotePrintedAt: "",

    calendarEventId: "",
    calendarEventUrl: "",
    calendarCreatedAt: "",

    manuallyEdited: false,
    lastEditedBy: "",
    lastEditedAt: "",

    createdAt: new Date(),
    updatedAt: new Date(),

    error: ""
  };
}

function validateBooking_(booking) {

  const errors = [];

  if (!booking.clientCompany)
    errors.push("Missing company");

  if (!booking.eventDate)
    errors.push("Missing event date");

  if (!booking.serviceTimes || booking.serviceTimes.length === 0)
    errors.push("Missing service time");

  if (!booking.pax)
    errors.push("Missing pax");

  if (!booking.location)
    errors.push("Missing location");

  if (!booking.items || booking.items.length === 0)
    errors.push("Missing line items");

  if (!booking.totalPrice || booking.totalPrice <= 0)
    errors.push("Missing total price");

  booking.validationErrors = errors;

  booking.status =
    errors.length > 0
      ? CONFIG.STATUS.NEEDS_REVIEW
      : CONFIG.STATUS.READY;

  return booking;
}