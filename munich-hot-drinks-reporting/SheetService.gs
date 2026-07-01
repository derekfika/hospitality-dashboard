function recordDrinkTap(payload) {
  setupHotDrinkTally();
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const settings = getSettings_();
    const floor = String(payload && payload.floor || "");
    const drink = String(payload && payload.drink || "");
    const clientTapId = String(payload && payload.clientTapId || "");
    const device = String(payload && payload.device || getUser_() || "Unknown device");

    if (settings.floors.indexOf(floor) === -1) throw new Error("Unknown floor: " + floor);
    if (settings.drinks.indexOf(drink) === -1) throw new Error("Unknown drink: " + drink);
    if (clientTapId && clientTapIdExists_(clientTapId)) {
      return { ok: true, duplicate: true, counts: getTodayCounts(), message: "Duplicate tap ignored." };
    }

    const now = new Date();
    const id = Utilities.getUuid();
    const row = [
      id,
      now,
      Utilities.formatDate(now, HOT_DRINKS_CONFIG.timezone, "yyyy-MM-dd"),
      Utilities.formatDate(now, HOT_DRINKS_CONFIG.timezone, "HH:mm:ss"),
      floor,
      drink,
      device,
      HOT_DRINKS_CONFIG.source,
      "ACTIVE",
      clientTapId
    ];
    const sheet = getSpreadsheet_().getSheetByName(HOT_DRINKS_CONFIG.sheets.drinkLog);
    sheet.appendRow(row);
    return { ok: true, id: id, counts: getTodayCounts() };
  } catch (error) {
    logAudit_("ERROR", "", payload && payload.floor, payload && payload.drink, getUser_(), error.message || String(error));
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function undoLastTap(payload) {
  setupHotDrinkTally();
  const floor = String(payload && payload.floor || "");
  const device = String(payload && payload.device || getUser_() || "Unknown device");
  const sheet = getSpreadsheet_().getSheetByName(HOT_DRINKS_CONFIG.sheets.drinkLog);
  const values = getLogRows_();
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    if (row.status !== "ACTIVE") continue;
    if (floor && row.floor !== floor) continue;
    const rowNumber = row.rowNumber;
    sheet.getRange(rowNumber, row.headerMap.Status).setValue("UNDONE");
    logAudit_("UNDO", row.id, row.floor, row.drink, device, "Marked row " + rowNumber + " as UNDONE.");
    return { ok: true, undone: row, counts: getTodayCounts() };
  }
  return { ok: false, message: "No active tap found to undo.", counts: getTodayCounts() };
}

function getTodayCounts() {
  setupHotDrinkTally();
  const settings = getSettings_();
  const today = Utilities.formatDate(new Date(), HOT_DRINKS_CONFIG.timezone, "yyyy-MM-dd");
  const counts = makeEmptyCounts_(settings);
  getLogRows_().forEach(function(row) {
    if (row.status !== "ACTIVE" || row.date !== today) return;
    counts.total += 1;
    if (!counts.byDrink[row.drink]) counts.byDrink[row.drink] = 0;
    if (!counts.byFloor[row.floor]) counts.byFloor[row.floor] = 0;
    counts.byDrink[row.drink] += 1;
    counts.byFloor[row.floor] += 1;
  });
  counts.date = today;
  return counts;
}

function clientTapIdExists_(clientTapId) {
  if (!clientTapId) return false;
  const sheet = getSpreadsheet_().getSheetByName(HOT_DRINKS_CONFIG.sheets.drinkLog);
  if (sheet.getLastRow() < 2) return false;
  const map = getHeaderMap_(sheet);
  const values = sheet.getRange(2, map["Client Tap ID"], sheet.getLastRow() - 1, 1).getDisplayValues();
  return values.some(function(row) { return row[0] === clientTapId; });
}

function getLogRows_() {
  const sheet = getOrCreateSheet_(getSpreadsheet_(), HOT_DRINKS_CONFIG.sheets.drinkLog, DRINK_LOG_HEADERS);
  const map = getHeaderMap_(sheet);
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.map(function(row, index) {
    return {
      rowNumber: index + 2,
      headerMap: map,
      id: row[map.ID - 1],
      timestamp: row[map.Timestamp - 1],
      date: dateKey_(row[map.Date - 1]),
      time: timeKey_(row[map.Time - 1]),
      floor: String(row[map.Floor - 1] || ""),
      drink: String(row[map.Drink - 1] || ""),
      device: String(row[map["Device/User"] - 1] || ""),
      source: String(row[map.Source - 1] || ""),
      status: String(row[map.Status - 1] || "ACTIVE")
    };
  });
}

function makeEmptyCounts_(settings) {
  const byDrink = {};
  const byFloor = {};
  settings.drinks.forEach(function(drink) { byDrink[drink] = 0; });
  settings.floors.forEach(function(floor) { byFloor[floor] = 0; });
  return { total: 0, byDrink: byDrink, byFloor: byFloor };
}

function logAudit_(action, submissionId, floor, drink, device, details) {
  const sheet = getOrCreateSheet_(getSpreadsheet_(), HOT_DRINKS_CONFIG.sheets.auditLog, AUDIT_LOG_HEADERS);
  sheet.appendRow([new Date(), action, submissionId || "", floor || "", drink || "", device || getUser_(), details || ""]);
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(header, index) { map[String(header).trim()] = index + 1; });
  return map;
}

function getUser_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "";
  } catch (error) {
    return "";
  }
}

function dateKey_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, HOT_DRINKS_CONFIG.timezone, "yyyy-MM-dd");
  }
  const text = String(value).trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return match[1] + "-" + match[2] + "-" + match[3];
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? text : Utilities.formatDate(parsed, HOT_DRINKS_CONFIG.timezone, "yyyy-MM-dd");
}

function timeKey_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, HOT_DRINKS_CONFIG.timezone, "HH:mm:ss");
  }
  return String(value).trim();
}
