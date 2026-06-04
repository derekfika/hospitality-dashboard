function testParseUploadedBookingFromMessageId() {
  const messageId = "19e692579772d596"
  const booking = buildBookingFromMessageId(messageId);
  writeBookingToSheet_(booking);
}

function isAngelCourtBookingForm_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();

  function norm(v) {
    return String(v || "").replace(/\s+/g, " ").trim().toLowerCase();
  }

  function hasText(text) {
    const target = norm(text);
    return values.some(row => row.some(cell => norm(cell) === target));
  }

  let hasItemHeader = false;

  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(norm);

    const hasItem = row.includes("item");
    const hasNumberRequired = row.some(v => v.includes("number required"));

    if (hasItem && hasNumberRequired) {
      hasItemHeader = true;
      break;
    }
  }

  return (
    hasItemHeader &&
    hasText("Company Name:") &&
    hasText("Date of event:") &&
    hasText("Total Number of people:") &&
    hasText("Grand Net Total")
  );
}

function buildBookingFromMessageId(messageId) {
  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  const atts = msg.getAttachments({ includeInlineImages: false });

  const xlsx = atts.find(att => isXlsx_(att));
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
        const fixedFallbackTime = normaliseHospitalityTime_(fallbackTime);

        booking.serviceTimes = [fixedFallbackTime];

        booking.items = booking.items.map(item => {
          if (!item.time) item.time = fixedFallbackTime;
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

    booking.hostEmail = extractEmailAddress_(msg.getFrom());
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

function normaliseBookingTimes_(booking) {
  booking.serviceTimes = (booking.serviceTimes || [])
    .map(forceLateHospitalityTimeToMorning_)
    .filter(Boolean);

  booking.items = (booking.items || []).map(item => {
    if (item.time) {
      item.time = forceLateHospitalityTimeToMorning_(item.time);
    }
    return item;
  });

  return booking;
}

function forceLateHospitalityTimeToMorning_(timeText) {
  const m = String(timeText || "").trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return timeText;

  let hour = Number(m[1]);
  const minute = m[2];

  if (hour >= 18 && hour <= 23) {
    hour -= 12;
  }

  return String(hour).padStart(2, "0") + ":" + minute;
}

function extractTimeFromText_(text) {
  const s = String(text || "");

  // 08:00pm / 08:00 pm / 8:00pm
  let m = s.match(/\b(\d{1,2}):(\d{2})\s*(am|pm)\b/i);
  if (m) {
    let hh = parseInt(m[1], 10);
    const mm = m[2];
    const ampm = m[3].toLowerCase();

    if (ampm === "am" && hh === 12) hh = 0;
    if (ampm === "pm" && hh < 12) hh += 12;

    return pad2_(hh) + ":" + mm;
  }

  // 08:00 / 8:30
  m = s.match(/\b(\d{1,2}):(\d{2})\b/);
  if (m) {
    return pad2_(m[1]) + ":" + m[2];
  }

  // 8am / 8 pm
  m = s.match(/\b(\d{1,2})\s*(am|pm)\b/i);
  if (m) {
    let hh = parseInt(m[1], 10);
    const ampm = m[2].toLowerCase();

    if (ampm === "am" && hh === 12) hh = 0;
    if (ampm === "pm" && hh < 12) hh += 12;

    return pad2_(hh) + ":00";
  }

  return "";
}

function isXlsx_(attachment) {
  const name = String(attachment.getName() || "").toLowerCase();
  const ct = String(attachment.getContentType() || "").toLowerCase();

  return (
    name.endsWith(".xlsx") ||
    ct === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
}

function extractEmailAddress_(fromText) {
  const text = String(fromText || "");

  let m = text.match(/<([^>]+)>/);
  if (m) return m[1].trim();

  m = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0].trim() : "";
}

function extractEmailName_(fromText) {
  const text = String(fromText || "").trim();

  let m = text.match(/^(.+?)\s*<.+>$/);
  if (m) return m[1].replace(/["']/g, "").trim();

  m = text.match(/^([^@]+)@/);
  if (m) {
    return m[1]
      .replace(/[._-]+/g, " ")
      .replace(/\b\w/g, s => s.toUpperCase())
      .trim();
  }

  return text;
}

function parseAngelCourtBookingSheet_(sheet) {
  let booking = createEmptyBooking_();

  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();
  const notes =
    getNotesBlockUnderLabel_(sheet, "Allergens & Dietary Requirements") ||
    getNotesBlockUnderLabel_(sheet, "NOTES:") ||
    getNotesBlockUnderLabel_(sheet, "Notes & dietary/allergen information") ||
    "";
  function norm(v) {
    return String(v || "").replace(/\s+/g, " ").trim().toLowerCase();
  }

  function findValueRightOfLabel(labelText) {
    const target = norm(labelText);

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (norm(values[r][c]) === target) {
          for (let offset = 1; offset <= 5; offset++) {
            const v = values[r][c + offset];
            if (v !== "" && v !== null && v !== undefined) return v;
          }
        }
      }
    }

    return "";
  }

  function findDisplayRightOfLabel(labelText) {
    const target = norm(labelText);

    for (let r = 0; r < displayValues.length; r++) {
      for (let c = 0; c < displayValues[r].length; c++) {
        if (norm(displayValues[r][c]) === target) {
          for (let offset = 1; offset <= 5; offset++) {
            const v = displayValues[r][c + offset];
            if (v !== "" && v !== null && v !== undefined) return String(v).trim();
          }
        }
      }
    }

    return "";
  }

  const company =
    findValueRightOfLabel("Company Name:") ||
    findValueRightOfLabel("Company name:");

  const pax =
    findValueRightOfLabel("Total Number of people:") ||
    findValueRightOfLabel("Number of guests:");

  const floorRaw =
    findValueRightOfLabel("Floor Level") ||
    findValueRightOfLabel("Floor level:");

  const hostName =
    findValueRightOfLabel("Name:") ||
    findValueRightOfLabel("Contact name:");

  const hostEmail =
    findValueRightOfLabel("Email:") ||
    findValueRightOfLabel("Contact email:");

  const dateRaw =
    findValueRightOfLabel("Date of event:") ||
    findDisplayRightOfLabel("Date of event:") ||
    findValueRightOfLabel("Date of delivery (DD/MM/YY):") ||
    findDisplayRightOfLabel("Date of delivery (DD/MM/YY):");

  const items = parseAngelCourtLineItems_(sheet);

  const serviceTimes = [...new Set(
    items
      .map(i => i.time)
      .filter(Boolean)
  )];

  let eventDate = parseDateToIsoDate_(dateRaw);

  if (!eventDate) {
    const fallbackDate = extractDateFromSheetName_(sheet.getName());
    if (fallbackDate) {
      eventDate = Utilities.formatDate(
        fallbackDate,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );
    }
  }

  const serviceTypes = [...new Set(
    items
      .map(i => i.section)
      .filter(Boolean)
      .map(s => String(s).trim())
  )];

  const totalPrice = findCurrencyValueByLabel_(sheet, "Grand Net Total");

  const mgmtFee = totalPrice * 0.08;
  const netPrice = totalPrice + mgmtFee;
  const vat = netPrice * 0.20;
  const grossPrice = netPrice + vat;

  booking.clientCompany = String(company || "").trim();
  booking.hostName = String(hostName || "").trim();
  booking.hostEmail = String(hostEmail || "").trim();
  booking.pax = pax || "";

  booking.eventDate = eventDate || "";

  booking.serviceType = serviceTypes.length
    ? serviceTypes.join(" / ")
    : "";

  booking.location = "One Angel Court";
  booking.floor = cleanFloor_(floorRaw);
  booking.notes = notes;

  booking.totalPrice = roundMoney_(totalPrice);
  booking.mgmtFee = roundMoney_(mgmtFee);
  booking.netPrice = roundMoney_(netPrice);
  booking.vat = roundMoney_(vat);
  booking.grossPrice = roundMoney_(grossPrice);

  booking.serviceTimes = (serviceTimes || [])
    .map(normaliseHospitalityTime_)
    .filter(Boolean);

  booking.items = (items || []).map(item => {
    if (item.time) {
      item.time =
        normaliseHospitalityTime_(item.time);
    }
    return item;
  });

  return booking;
}

function normaliseHospitalityTime_(timeText) {
  console.log("NORMALISE INPUT:", timeText);
  if (!timeText) return "";

  let t = String(timeText).trim().toLowerCase();

  const match = t.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/);
  if (!match) return timeText;

  let hour = Number(match[1]);
  const minute = match[2] || "00";
  const meridian = match[3] || "";

  if (meridian === "pm" && hour !== 12) {
    hour += 12;
  }

  if (meridian === "am" && hour === 12) {
    hour = 0;
  }

  // Angel Court hospitality sanity check:
  // 20:00, 23:30 etc. almost always means 08:00, 11:30.
  if (hour >= 18 && hour <= 23) {
    hour -= 12;
  }
  console.log("NORMALISE OUTPUT:", `${String(hour).padStart(2, "0")}:${minute}`);
  return `${String(hour).padStart(2, "0")}:${minute}`;
}

function parseRequiredQty_(value) {
  if (value === "" || value === null || value === undefined) return null;

  if (typeof value === "number") {
    return value > 0 ? value : null;
  }

  const text = String(value).trim();

  if (!text) return null;

  const match = text.match(/\d+(\.\d+)?/);
  if (!match) return null;

  const n = Number(match[0]);
  return n > 0 ? n : null;
}

function parseAngelCourtLineItems_(sheet) {
  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();

  function norm(v) {
    return String(v || "").replace(/\s+/g, " ").trim().toLowerCase();
  }

  let headerRow = -1;
  let cols = {
    item: -1,
    info: -1,
    qty: -1,
    time: -1,
    comment: -1
  };

  for (let r = 0; r < values.length; r++) {
    const row = displayValues[r].map(norm);
    const itemCol = row.findIndex(v =>
      v === "item" ||
      v.includes("item")
    );

    const qtyCol = row.findIndex(v =>
      v.includes("number required") ||
      v.includes("number") ||
      v.includes("qty") ||
      v.includes("quantity") ||
      v.includes("amount required")
    );

    if (itemCol !== -1 && qtyCol !== -1) {
      headerRow = r;
      cols.item = itemCol;
      cols.qty = qtyCol;

      cols.info = row.indexOf("info");

      cols.time = row.findIndex(v =>
        v.includes("time required") ||
        v.includes("service time") ||
        v === "time"
      );

      cols.comment = row.findIndex(v =>
        v.includes("comment") ||
        v.includes("notes")
      );

      break;
    }
  }

  if (headerRow === -1) return [];

  const fallbackTime = extractTimeFromSheetName_(sheet.getName());

  const items = [];
  let currentSection = "";
  let emptyRun = 0;

  for (let r = headerRow + 1; r < values.length; r++) {
    const itemRaw = cols.item >= 0 ? values[r][cols.item] : "";
    const infoRaw = cols.info >= 0 ? values[r][cols.info] : "";
    const qtyRaw = cols.qty >= 0 ? values[r][cols.qty] : "";
    const timeRaw = cols.time >= 0 ? displayValues[r][cols.time] : "";
    const commentRaw = cols.comment >= 0 ? values[r][cols.comment] : "";

    const hasAny =
      itemRaw !== "" && itemRaw !== null ||
      infoRaw !== "" && infoRaw !== null ||
      qtyRaw !== "" && qtyRaw !== null ||
      commentRaw !== "" && commentRaw !== null;

    if (!hasAny) {
      emptyRun++;
      if (emptyRun >= 15) break;
      continue;
    }

    emptyRun = 0;

    const itemText = String(itemRaw || "").trim();
    const infoText = String(infoRaw || "").trim();
    const commentText = String(commentRaw || "").trim();

    const qtyNum = parseRequiredQty_(qtyRaw);

    const isSection =
      itemText &&
      !infoText &&
      !commentText &&
      (!qtyNum || qtyNum <= 0);

    if (isSection) {
      const lowerItem = itemText.toLowerCase();

      if (
        lowerItem.includes("please note") ||
        lowerItem.includes("quotes are subject") ||
        lowerItem.includes("notice") ||
        lowerItem.includes("cancellation")
      ) {
        break;
      }

      currentSection = itemText;
      continue;
    }

    if (!qtyNum || qtyNum <= 0) continue;

    const split = splitOrderNameAndDetail_(itemText);

    const timeText =
      normaliseHospitalityTime_(timeRaw) ||
      normaliseHospitalityTime_(fallbackTime) ||
      "";

    items.push({
      section: currentSection,
      time: timeText,
      name: split.name,
      detail: split.detail,
      info: infoText,
      qty: qtyNum % 1 === 0 ? String(qtyNum.toFixed(0)) : String(qtyNum),
      comment: commentText
    });
  }

  Logger.log("ITEM COUNT = " + items.length);

  if (items.length) {
    Logger.log("FIRST ITEM = " + JSON.stringify(items[0]));
  }

  return items;
}

function findLikelyQtyInRow_(valueRow, displayRow, itemCol) {
  for (let c = 0; c < valueRow.length; c++) {
    if (c === itemCol) continue;

    const raw = valueRow[c];
    const display = String(displayRow[c] || "").trim();

    if (typeof raw === "number" && raw > 0 && raw < 1000) {
      return raw;
    }

    if (/^\d+(\.\d+)?$/.test(display)) {
      const n = Number(display);
      if (n > 0 && n < 1000) return n;
    }
  }

  return "";
}

function cleanFloor_(value) {
  const text = String(value || "").trim();
  return text.replace(/^floor\s+/i, "").trim();
}

function roundMoney_(n) {
  return Math.round((Number(n || 0) + Number.EPSILON) * 100) / 100;
}

function findCurrencyValueByLabel_(sheet, labelText) {
  const values = sheet.getDataRange().getValues();
  const displays = sheet.getDataRange().getDisplayValues();

  function norm(v) {
    return String(v || "").replace(/\s+/g, " ").trim().toLowerCase();
  }

  const target = norm(labelText);

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (norm(values[r][c]) === target || norm(displays[r][c]) === target) {
        for (let offset = 1; offset <= 6; offset++) {
          const raw = values[r][c + offset];
          const display = displays[r][c + offset];

          if (typeof raw === "number" && isFinite(raw)) return raw;

          const cleaned = String(display || "")
            .replace(/[£,]/g, "")
            .trim();

          const n = parseFloat(cleaned);
          if (isFinite(n)) return n;
        }
      }
    }
  }

  return 0;
}

function getNotesBlockUnderLabel_(sheet, labelText) {
  const values = sheet.getDataRange().getDisplayValues();

  function norm(v) {
    return String(v || "").replace(/\s+/g, " ").trim().toLowerCase();
  }

  const target = norm(labelText);

  let labelRow = -1;
  let labelCol = -1;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (norm(values[r][c]) === target) {
        labelRow = r;
        labelCol = c;
        break;
      }
    }
    if (labelRow !== -1) break;
  }

  if (labelRow === -1) return "";

  const lines = [];
  let blankRun = 0;

  // Search a wider block below the label because this section uses merged/spaced cells.
  const startRow = labelRow + 1;
  const maxRowsToScan = 30;

  for (let r = startRow; r < Math.min(values.length, startRow + maxRowsToScan); r++) {
    const rowTexts = [];

    for (let c = 0; c < values[r].length; c++) {
      const t = String(values[r][c] || "").trim();
      if (t) rowTexts.push(t);
    }

    if (rowTexts.length === 0) {
      blankRun++;
      if (blankRun >= 8 && lines.length > 0) break;
      continue;
    }

    blankRun = 0;

    const joined = rowTexts.join(" ").replace(/\s+/g, " ").trim();
    const lower = joined.toLowerCase();

    // Stop if we hit another obvious section/financial area.
    if (
      lower.includes("notice") ||
      lower.includes("cancellation") ||
      lower.includes("grand net total") ||
      lower.includes("please note that all quotes")
    ) {
      break;
    }

    // Avoid accidentally repeating the label.
    if (norm(joined) === target) continue;

    lines.push(joined);
  }

  const cleaned = [];
  for (const line of lines) {
    if (cleaned.length === 0 || cleaned[cleaned.length - 1] !== line) {
      cleaned.push(line);
    }
  }

  return cleaned.join("\n");
}

function parseDateToIsoDate_(raw) {
  const d = parseFlexibleDateTime_(raw);
  if (!d) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function parseFlexibleDateTime_(raw) {
  if (!raw) return null;

  if (Object.prototype.toString.call(raw) === "[object Date]" && !isNaN(raw.getTime())) {
    return raw;
  }

  const text = String(raw)
    .replace(/\r?\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!text) return null;

  let m;

  m = text.match(/^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?$/);
  if (m) {
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    let yyyy = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
    if (yyyy < 100) yyyy += 2000;

    const d = new Date(yyyy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  m = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    let yyyy = parseInt(m[3], 10);
    if (yyyy < 100) yyyy += 2000;

    const d = new Date(yyyy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  return null;
}

function extractDateFromSheetName_(name) {
  const text = String(name || "");

  const m = text.match(/\b(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\b/);
  if (!m) return null;

  const dd = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  let yyyy = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
  if (yyyy < 100) yyyy += 2000;

  const d = new Date(yyyy, mm - 1, dd);
  return isNaN(d.getTime()) ? null : d;
}

function extractTimeFromSheetName_(name) {
  const text = String(name || "");

  let m = text.match(/\b(\d{1,2})[:.](\d{2})\b/);
  if (m) return pad2_(m[1]) + ":" + m[2];

  m = text.match(/\b(\d{1,2})\s*(am|pm)\b/i);
  if (m) {
    let hh = parseInt(m[1], 10);
    const ampm = m[2].toLowerCase();

    if (ampm === "am" && hh === 12) hh = 0;
    if (ampm === "pm" && hh < 12) hh += 12;

    return pad2_(hh) + ":00";
  }

  return "";
}

function normaliseTimeText_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }

  const text = String(value).trim();

  let m = text.match(/^(\d{1,2})[:.](\d{2})$/);
  if (m) return pad2_(m[1]) + ":" + m[2];

  m = text.match(/^(\d{1,2})\s*(am|pm)$/i);
  if (m) {
    let hh = parseInt(m[1], 10);
    const ampm = m[2].toLowerCase();

    if (ampm === "am" && hh === 12) hh = 0;
    if (ampm === "pm" && hh < 12) hh += 12;

    return pad2_(hh) + ":00";
  }

  return text;
}

function pad2_(n) {
  return String(n).padStart(2, "0");
}

function splitOrderNameAndDetail_(text) {
  const s = String(text || "").trim();
  const parts = s.split(/\s+-\s+/);

  if (parts.length <= 1) {
    return { name: s, detail: "" };
  }

  return {
    name: parts.shift().trim(),
    detail: parts.join(" - ").trim()
  };
}

function debugDashboardCount() {
  const bookings = getDashboardBookings();
  Logger.log("Bookings found: " + bookings.length);
  Logger.log(JSON.stringify(bookings.map(b => b.BookingID)));
}