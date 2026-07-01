function getDashboardReport(filters) {
  setupHotDrinkTally();
  const settings = getSettings_();
  const normalized = normalizeFilters_(filters || {});
  const rows = getLogRows_().filter(function(row) {
    return row.status === "ACTIVE" && row.date >= normalized.startDate && row.date <= normalized.endDate;
  }).filter(function(row) {
    if (normalized.floor !== "Combined" && row.floor !== normalized.floor) return false;
    if (normalized.drink !== "All" && row.drink !== normalized.drink) return false;
    if (normalized.excludeBankHolidays && settings.bankHolidays.indexOf(row.date) !== -1) return false;
    const weekday = weekdayName_(row.date);
    if (normalized.weekdays.length && normalized.weekdays.indexOf(weekday) === -1) return false;
    return true;
  });

  const summary = summarizeRows_(rows, normalized, settings);
  cacheDashboardReport_(normalized, summary);
  return { ok: true, filters: normalized, settings: settings, summary: summary };
}

function refreshDashboardData() {
  return getDashboardReport({ preset: "today" });
}

function exportFilteredCsv(filters) {
  const report = getDashboardReport(filters);
  const normalized = report.filters;
  const rows = getLogRows_().filter(function(row) {
    return row.status === "ACTIVE" && row.date >= normalized.startDate && row.date <= normalized.endDate;
  }).filter(function(row) {
    return (normalized.floor === "Combined" || row.floor === normalized.floor) &&
      (normalized.drink === "All" || row.drink === normalized.drink);
  });
  const header = ["ID", "Date", "Time", "Floor", "Drink", "Device/User", "Status"];
  const csvRows = [header].concat(rows.map(function(row) {
    return [row.id, row.date, row.time, row.floor, row.drink, row.device, row.status];
  }));
  return {
    ok: true,
    filename: "munich-re-hot-drinks-" + normalized.startDate + "-to-" + normalized.endDate + ".csv",
    csv: csvRows.map(function(row) { return row.map(csvCell_).join(","); }).join("\n")
  };
}

function exportPdfReport(filters) {
  const report = getDashboardReport(filters);
  const normalized = report.filters;
  const summary = report.summary;
  const template = HtmlService.createTemplateFromFile("PdfReport");
  template.report = buildPdfReportModel_(normalized, summary);
  const filename = "Munich-Re-hot-drink-report-" + normalized.startDate + "-to-" + normalized.endDate + ".pdf";
  const blob = template.evaluate().getBlob().getAs(MimeType.PDF).setName(filename);
  return {
    ok: true,
    filename: filename,
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

function clearTestData(confirmText) {
  if (String(confirmText || "") !== "CLEAR TEST DATA") throw new Error("Type CLEAR TEST DATA to confirm.");
  const sheet = getSpreadsheet_().getSheetByName(HOT_DRINKS_CONFIG.sheets.drinkLog);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, cleared: 0 };
  const map = getHeaderMap_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  let cleared = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    const source = String(values[i][map.Source - 1] || "");
    const device = String(values[i][map["Device/User"] - 1] || "");
    if (source.indexOf("TEST") !== -1 || device.indexOf("TEST") !== -1) {
      sheet.deleteRow(i + 2);
      cleared += 1;
    }
  }
  logAudit_("CLEAR_TEST_DATA", "", "", "", getUser_(), cleared + " test rows deleted.");
  return { ok: true, cleared: cleared };
}

function normalizeFilters_(filters) {
  const today = new Date();
  const preset = String(filters.preset || "today");
  let start = filters.startDate ? parseLocalDate_(filters.startDate) : today;
  let end = filters.endDate ? parseLocalDate_(filters.endDate) : today;
  const day = today.getDay();
  if (preset === "yesterday") start = end = addDays_(today, -1);
  if (preset === "thisWeek") {
    start = addDays_(today, day === 0 ? -6 : 1 - day);
    end = today;
  }
  if (preset === "lastWeek") {
    const thisMonday = addDays_(today, day === 0 ? -6 : 1 - day);
    start = addDays_(thisMonday, -7);
    end = addDays_(thisMonday, -1);
  }
  if (preset === "thisMonth") {
    start = new Date(today.getFullYear(), today.getMonth(), 1);
    end = today;
  }
  if (preset === "lastMonth") {
    start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    end = new Date(today.getFullYear(), today.getMonth(), 0);
  }
  let weekdays = (filters.weekdays || []).filter(Boolean);
  if (preset === "tueThu") weekdays = ["Tuesday", "Wednesday", "Thursday"];
  if (preset === "monThu") weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday"];
  if (preset === "excludeFridays") weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Saturday", "Sunday"];
  return {
    preset: preset,
    startDate: dateKey_(start),
    endDate: dateKey_(end),
    floor: filters.floor || "Combined",
    drink: filters.drink || "All",
    weekdays: weekdays,
    excludeBankHolidays: filters.excludeBankHolidays !== false,
    excludeClosedPeriods: filters.excludeClosedPeriods !== false
  };
}

function summarizeRows_(rows, filters, settings) {
  const byDrink = {};
  const byFloor = {};
  const byWeekday = {};
  const byDate = {};
  const byHour = {};
  const heatmap = {};
  settings.drinks.forEach(function(drink) { byDrink[drink] = 0; });
  settings.floors.forEach(function(floor) { byFloor[floor] = 0; });
  ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"].forEach(function(day) {
    byWeekday[day] = 0;
    heatmap[day] = {};
    HOT_DRINKS_CONFIG.timeBuckets.forEach(function(bucket) { heatmap[day][bucket.label] = 0; });
  });

  rows.forEach(function(row) {
    byDrink[row.drink] = (byDrink[row.drink] || 0) + 1;
    byFloor[row.floor] = (byFloor[row.floor] || 0) + 1;
    byDate[row.date] = (byDate[row.date] || 0) + 1;
    const weekday = weekdayName_(row.date);
    byWeekday[weekday] = (byWeekday[weekday] || 0) + 1;
    const hour = row.time.slice(0, 2) + ":00";
    byHour[hour] = (byHour[hour] || 0) + 1;
    const bucket = bucketForTime_(row.time);
    if (bucket && (!filters.excludeClosedPeriods || !bucket.closed)) heatmap[weekday][bucket.label] += 1;
  });

  const activeDates = Object.keys(byDate);
  const drinkMix = {};
  Object.keys(byDrink).forEach(function(drink) {
    drinkMix[drink] = rows.length ? Math.round((byDrink[drink] / rows.length) * 1000) / 10 : 0;
  });

  return {
    total: rows.length,
    byDrink: byDrink,
    byFloor: byFloor,
    byWeekday: byWeekday,
    byDate: byDate,
    byHour: byHour,
    averagePerDay: activeDates.length ? round_(rows.length / activeDates.length) : 0,
    averageByWeekday: averageByWeekday_(byDate),
    tuesdayThursdayDailyAverage: averageForWeekdays_(byDate, ["Tuesday", "Wednesday", "Thursday"]),
    mondayThursdayDailyAverage: averageForWeekdays_(byDate, ["Monday", "Tuesday", "Wednesday", "Thursday"]),
    averageByFloor: averageMap_(byFloor, activeDates.length),
    averageByDrink: averageMap_(byDrink, activeDates.length),
    peakHour: maxKey_(byHour),
    quietestHour: minKey_(byHour),
    busiestDay: maxKey_(byDate),
    heatmap: heatmap,
    drinkMix: drinkMix,
    trend: activeDates.sort().map(function(date) { return { date: date, total: byDate[date] }; })
  };
}

function cacheDashboardReport_(filters, summary) {
  const sheet = getOrCreateSheet_(getSpreadsheet_(), HOT_DRINKS_CONFIG.sheets.dashboardData, DASHBOARD_DATA_HEADERS);
  sheet.appendRow([new Date(), filters.startDate, filters.endDate, JSON.stringify(filters), JSON.stringify(summary)]);
}

function averageByWeekday_(byDate) {
  const totals = {};
  const counts = {};
  Object.keys(byDate).forEach(function(date) {
    const weekday = weekdayName_(date);
    totals[weekday] = (totals[weekday] || 0) + byDate[date];
    counts[weekday] = (counts[weekday] || 0) + 1;
  });
  Object.keys(totals).forEach(function(day) { totals[day] = round_(totals[day] / counts[day]); });
  return totals;
}

function averageForWeekdays_(byDate, weekdays) {
  const dates = Object.keys(byDate).filter(function(date) { return weekdays.indexOf(weekdayName_(date)) !== -1; });
  const total = dates.reduce(function(sum, date) { return sum + byDate[date]; }, 0);
  return dates.length ? round_(total / dates.length) : 0;
}

function averageMap_(map, divisor) {
  const output = {};
  Object.keys(map).forEach(function(key) { output[key] = divisor ? round_(map[key] / divisor) : 0; });
  return output;
}

function bucketForTime_(time) {
  return HOT_DRINKS_CONFIG.timeBuckets.find(function(bucket) {
    return time >= bucket.start + ":00" && time < bucket.end + ":00";
  }) || null;
}

function weekdayName_(dateString) {
  const date = parseLocalDate_(dateString);
  return ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"][date.getDay()];
}

function parseLocalDate_(dateString) {
  const parts = String(dateString || dateKey_(new Date())).split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function addDays_(date, days) {
  const output = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  output.setDate(output.getDate() + days);
  return output;
}

function maxKey_(map) {
  const keys = Object.keys(map);
  if (!keys.length) return "";
  return keys.reduce(function(best, key) { return map[key] > map[best] ? key : best; }, keys[0]);
}

function minKey_(map) {
  const keys = Object.keys(map).filter(function(key) { return map[key] > 0; });
  if (!keys.length) return "";
  return keys.reduce(function(best, key) { return map[key] < map[best] ? key : best; }, keys[0]);
}

function round_(value) {
  return Math.round(Number(value || 0) * 10) / 10;
}

function csvCell_(value) {
  const text = String(value == null ? "" : value);
  return /[",\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
}

function buildPdfReportModel_(filters, summary) {
  const generatedAt = Utilities.formatDate(new Date(), HOT_DRINKS_CONFIG.timezone, "d MMMM yyyy HH:mm");
  const activeDays = Object.keys(summary.byDate || {}).length;
  const drinkRows = mapRows_(summary.byDrink || {}, summary.total, summary.drinkMix || {});
  const floorRows = mapRows_(summary.byFloor || {}, summary.total, null);
  const weekdayRows = mapRows_(summary.byWeekday || {}, summary.total, null);
  const trend = summary.trend || [];
  const trendMax = Math.max(1, Math.max.apply(null, trend.map(function(item) { return item.total; }).concat([0])));
  const heatmap = summary.heatmap || {};
  const heatmapMax = maxNestedValue_(heatmap);

  return {
    title: "Hot drink usage report",
    strapline: "Not if, but how",
    period: formatDisplayDate_(filters.startDate) + " - " + formatDisplayDate_(filters.endDate),
    generatedAt: generatedAt,
    filters: {
      floor: filters.floor,
      drink: filters.drink,
      preset: filters.preset,
      bankHolidays: filters.excludeBankHolidays ? "Excluded" : "Included",
      closedPeriods: filters.excludeClosedPeriods ? "Excluded" : "Included"
    },
    kpis: [
      { label: "Total drinks", value: summary.total || 0 },
      { label: "Average per active day", value: summary.averagePerDay || 0 },
      { label: "Peak hour", value: summary.peakHour || "-" },
      { label: "Busiest day", value: formatDisplayDate_(summary.busiestDay) || "-" },
      { label: "Tue-Thu daily average", value: summary.tuesdayThursdayDailyAverage || 0 },
      { label: "Mon-Thu daily average", value: summary.mondayThursdayDailyAverage || 0 }
    ],
    activeDays: activeDays,
    quietestHour: summary.quietestHour || "-",
    drinkRows: drinkRows,
    floorRows: floorRows,
    weekdayRows: weekdayRows,
    trendRows: trend.map(function(item) {
      return {
        date: formatShortDate_(item.date),
        total: item.total,
        height: Math.max(4, Math.round(item.total / trendMax * 100))
      };
    }),
    heatmapDays: Object.keys(heatmap),
    heatmapBuckets: HOT_DRINKS_CONFIG.timeBuckets.map(function(bucket) { return bucket.label; }),
    heatmap: heatmap,
    heatmapMax: heatmapMax,
    logos: getReportLogos_()
  };
}

function mapRows_(map, total, percentages) {
  const max = Math.max(1, Math.max.apply(null, Object.keys(map).map(function(key) { return Number(map[key] || 0); }).concat([0])));
  return Object.keys(map).map(function(key) {
    const value = Number(map[key] || 0);
    const percent = percentages && percentages[key] != null
      ? percentages[key]
      : total ? Math.round(value / total * 1000) / 10 : 0;
    return {
      label: key,
      value: value,
      percent: percent,
      width: Math.max(2, Math.round(value / max * 100))
    };
  });
}

function maxNestedValue_(map) {
  let max = 1;
  Object.keys(map).forEach(function(day) {
    Object.keys(map[day] || {}).forEach(function(bucket) {
      max = Math.max(max, Number(map[day][bucket] || 0));
    });
  });
  return max;
}

function heatOpacity_(value, max) {
  const number = Number(value || 0);
  if (!number) return 0.04;
  return Math.max(0.12, Math.min(0.9, number / Math.max(1, max)));
}

function formatDisplayDate_(dateString) {
  if (!dateString) return "";
  const date = parseLocalDate_(dateString);
  return Utilities.formatDate(date, HOT_DRINKS_CONFIG.timezone, "d MMM yyyy");
}

function formatShortDate_(dateString) {
  if (!dateString) return "";
  const date = parseLocalDate_(dateString);
  return Utilities.formatDate(date, HOT_DRINKS_CONFIG.timezone, "d MMM");
}

function getReportLogos_() {
  return {
    munichRe: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 560 120" role="img" aria-label="Munich Re logo"><g fill="#00538A"><path d="M28 18h70l28 42-28 42H28L0 60 28 18Zm8 12L16 60l20 30h56l20-30-20-30H36Z"/><path d="M21 48h86v8H21zM21 64h86v8H21z"/><text x="155" y="78" font-family="Arial, Helvetica, sans-serif" font-size="58" font-weight="700" letter-spacing="-2">Munich Re</text></g></svg>',
    fika: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1726 740" role="img" aria-label="Fika logo"><path fill="#4F34C7" fill-rule="evenodd" d="M429 0L0 0L0 729L132 729L133 428L395 428L395 296L132 295L132 133L429 132ZM492 194L492 729L624 729L624 194ZM492 0L492 132L624 132L624 0ZM687 0L687 729L820 729L937 573L1054 729L1221 729L1020 462L1221 194L1055 194L820 507L819 0ZM1259 266L1241 286L1222 312L1213 327L1197 361L1187 392L1181 422L1179 441L1179 483L1183 514L1189 539L1203 577L1216 602L1231 625L1247 645L1273 671L1308 697L1346 717L1380 729L1419 737L1441 739L1469 739L1489 737L1518 731L1539 724L1559 715L1584 700L1599 688L1601 692L1603 729L1725 729L1725 194L1603 194L1601 232L1599 236L1587 226L1563 211L1544 202L1518 193L1496 188L1470 185L1440 185L1412 188L1373 197L1346 207L1325 217L1293 237ZM1450 317L1473 318L1500 324L1525 335L1547 350L1568 371L1585 397L1594 419L1600 449L1600 475L1593 508L1581 534L1567 554L1547 574L1530 586L1515 594L1489 603L1458 607L1432 605L1406 598L1383 587L1362 572L1345 555L1329 532L1317 504L1311 473L1312 442L1317 420L1329 392L1345 369L1368 347L1396 330L1422 321Z"/></svg>',
    bloom: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 557 156" role="img" aria-label="Bloom logo"><g fill="none" stroke="#05090B" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4v148" /><path d="M30 5h25c16 0 29 15 29 34s-13 35-29 35H30" /><path d="M30 152h37c21 0 37-17 37-38 0-16-9-30-22-35" /><path d="M134 4v148h56" /><path d="M199 66V51c0-27 20-47 47-47s47 20 47 47v15" /><path d="M199 90v15c0 27 20 47 47 47s47-20 47-47V90" /><path d="M322 66V51c0-27 20-47 47-47s47 20 47 47v15" /><path d="M322 90v15c0 27 20 47 47 47s47-20 47-47V90" /><path d="M444 152V4l57 76" /><path d="M553 4v148" /><path d="M553 4l-40 56" /></g></svg>'
  };
}
