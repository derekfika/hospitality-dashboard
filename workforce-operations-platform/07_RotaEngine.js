function getRotaAppOptions() {
  const spreadsheet = getWorkforceSpreadsheet_();
  const sites = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.sites)
  ).filter(function(site) {
    return site["Site ID"] && String(site.Active).toUpperCase() !== "FALSE";
  }).map(function(site) {
    return {
      siteId: String(site["Site ID"]),
      siteName: String(site["Site Name"] || site["Site ID"])
    };
  }).sort(function(a, b) {
    return a.siteName.localeCompare(b.siteName);
  });
  return {
    ok: true,
    sites: sites,
    defaultWeekStart: getNextMonday_()
  };
}

function getWeeklyRotaBoard(siteId, weekStartDate) {
  const spreadsheet = getWorkforceSpreadsheet_();
  const cleanSiteId = String(siteId || "").trim();
  if (!cleanSiteId) throw new Error("Choose a site first.");
  const weekStart = normaliseWeekStart_(weekStartDate);
  const weekDates = getWeekDates_(weekStart);
  const templates = uniqueWorkforceRotaTemplates_(readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.rotaTemplates)
  ).filter(function(template) {
    return String(template["Site ID"] || "") === cleanSiteId &&
      workforceBoolean_(template.Active);
  }));
  if (!templates.length) {
    return {
      ok: true,
      siteId: cleanSiteId,
      siteName: cleanSiteId,
      weekStart: weekStart,
      days: [],
      totalRows: 0,
      totalGaps: 0,
      message: "No standard rota templates found for this site."
    };
  }
  const absences = buildAbsenceIndex_(readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.absences)
  ));
  const rotaExceptions = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.rotaExceptions)
  );
  const exceptions = buildRotaExceptionIndex_(rotaExceptions);
  const coverIndex = buildRotaCoverIndex_(rotaExceptions);
  const gapsByKey = buildCoverageGapIndex_(readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.coverageGaps)
  ));

  const siteName = (templates[0] && templates[0]["Site Name"]) || cleanSiteId;
  const days = weekDates.map(function(day) {
    const rows = templates
      .filter(function(template) {
        return String(template.Weekday || "") === day.weekday &&
          String(template["Standard Status"] || "").toUpperCase() !== "OFF";
      })
      .map(function(template) {
        const person = String(template["Employee Name"] || "").trim();
        const exception = exceptions[getGapTemplateDateKey_(template, day.date)];
        const cover = coverIndex[getCoverDateRoleKey_(template["Site ID"], day.date, template.Role)];
        const absence = absences[getGapPersonDateKey_(person, day.date)];
        const gap = gapsByKey[getGapTemplateDateKey_(template, day.date)];
        const status = exception
          ? String(exception["Actual Status"] || exception["Exception Type"] || "Exception")
          : cover && absence
            ? "Covered"
          : absence
            ? String(absence["Absence Type"] || "Absent")
            : String(template["Standard Status"] || "IN");
        return {
          role: String(template.Role || ""),
          employeeName: person,
          standardStatus: String(template["Standard Status"] || ""),
          status: status,
          exceptionType: exception ? String(exception["Exception Type"] || "") : "",
          gapType: gap ? String(gap["Gap Type"] || "") : "",
          coverName: cover ? String(cover["Employee Name"] || "") : "",
          coverType: cover ? String(cover["Exception Type"] || "Cover") : "",
          hasGap: Boolean((gap || absence || (exception && !isGapClearedByException_(exception))) && !cover),
          priority: gap ? Number(gap.Priority || 3) : getGapPriority_(template.Role),
          notes: exception ? String(exception.Notes || "") : ""
        };
      })
      .sort(function(a, b) {
        return a.priority - b.priority || a.role.localeCompare(b.role);
      });
    return {
      date: day.date,
      weekday: day.weekday,
      rows: rows,
      gapCount: rows.filter(function(row) { return row.hasGap; }).length
    };
  });

  return {
    ok: true,
    siteId: cleanSiteId,
    siteName: siteName,
    weekStart: weekStart,
    days: days,
    totalRows: days.reduce(function(total, day) { return total + day.rows.length; }, 0),
    totalGaps: days.reduce(function(total, day) { return total + day.gapCount; }, 0)
  };
}

function generateReliefSuggestions(daysAhead) {
  detectCoverageGaps(daysAhead || 28);
  const spreadsheet = getWorkforceSpreadsheet_();
  const gaps = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.coverageGaps)
  ).filter(function(gap) {
    return String(gap.Status || "").toLowerCase() !== "resolved";
  });
  const absences = buildAbsenceIndex_(readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.absences)
  ));
  const reliefAvailability = buildReliefAvailabilityIndex_(readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.reliefAvailability)
  ));
  const coverHistory = buildCoverHistoryIndex_(readCoverHistoryRows_(spreadsheet));
  const staff = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory)
  ).filter(function(person) {
    return isReliefCandidate_(person, reliefAvailability);
  });
  const suggestions = [];

  gaps.forEach(function(gap) {
    const availableRelief = reliefAvailability[normaliseWorkforceDate_(gap.Date)] || [];
    const candidates = staff
      .filter(function(person) {
        return String(person.Name || "") &&
          normaliseWorkforcePerson_(person.Name) !== normaliseWorkforcePerson_(gap["Employee Name"]) &&
          !absences[getGapPersonDateKey_(person.Name, normaliseWorkforceDate_(gap.Date))];
      })
      .map(function(person) {
        const reliefMatch = findReliefAvailabilityForPerson_(availableRelief, person.Name);
        const historySignal = getCoverHistorySignal_(coverHistory, gap, person);
        return {
          person: person,
          reliefMatch: reliefMatch,
          historySignal: historySignal,
          score: scoreReliefCandidate_(person, gap, historySignal) +
            (reliefMatch ? 60 : 0)
        };
      })
      .sort(function(a, b) { return b.score - a.score; })
      .slice(0, 3);
    candidates.forEach(function(candidate, index) {
      const id = [
        "relief",
        gap["Gap ID"],
        slugifyWorkforce_(candidate.person.Name),
        index + 1
      ].join("_");
      suggestions.push({
        "Suggestion ID": id,
        "Gap ID": gap["Gap ID"],
        "Site ID": gap["Site ID"],
        "Date": normaliseWorkforceDate_(gap.Date),
        "Role": gap.Role,
          "Suggested Employee ID": candidate.person["Employee ID"],
          "Suggested Employee Name": candidate.person.Name,
          "Reason": buildReliefReason_(candidate.person, gap, candidate.reliefMatch, candidate.historySignal),
        "Score": candidate.score,
        "Reviewed": false,
        "Approved": false
      });
    });
  });

  const sheet = ensureReliefSuggestionsSheet_(spreadsheet);
  replaceWorkforceSheetRowsFast_(sheet, suggestions);
  return {
    ok: true,
    gapsReviewed: gaps.length,
    suggestionsCreated: suggestions.length
  };
}

function isReliefCandidate_(person, reliefAvailability) {
  if (String(person["Employment Status"] || "").toLowerCase() === "terminated") return false;
  if (workforceBoolean_(person.Terminated)) return false;
  if (workforceBoolean_(person["Relief Team"])) return true;
  const personName = normaliseWorkforcePerson_(person.Name);
  return Object.keys(reliefAvailability || {}).some(function(date) {
    return (reliefAvailability[date] || []).some(function(row) {
      return normaliseWorkforcePerson_(row["Relief Name"]) === personName;
    });
  });
}

function getReliefSuggestionSummary(limit) {
  const spreadsheet = getWorkforceSpreadsheet_();
  const suggestions = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.reliefSuggestions)
  ).sort(function(a, b) {
    return String(a.Date || "").localeCompare(String(b.Date || "")) ||
      Number(b.Score || 0) - Number(a.Score || 0);
  });
  return {
    ok: true,
    count: suggestions.length,
    suggestions: suggestions.slice(0, Math.max(1, Number(limit || 12)))
  };
}

function generateReliefRotaAssignments(daysAhead) {
  const suggestionResult = generateReliefSuggestions(daysAhead || 28);
  const spreadsheet = getWorkforceSpreadsheet_();
  const gapsById = {};
  readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.coverageGaps))
    .filter(function(gap) {
      return String(gap.Status || "").toLowerCase() !== "resolved";
    })
    .forEach(function(gap) {
      gapsById[String(gap["Gap ID"] || "")] = gap;
    });
  const suggestions = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.reliefSuggestions)
  ).sort(function(a, b) {
    return String(a.Date || "").localeCompare(String(b.Date || "")) ||
      Number(b.Score || 0) - Number(a.Score || 0);
  });
  const staffByName = {};
  readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory))
    .forEach(function(person) {
      staffByName[normaliseWorkforcePerson_(person.Name)] = person;
    });
  const usedByDate = {};
  const assignedGaps = {};
  const now = new Date();
  const assignments = [];
  suggestions.forEach(function(suggestion) {
    const gapId = String(suggestion["Gap ID"] || "");
    const gap = gapsById[gapId];
    if (!gap || assignedGaps[gapId]) return;
    const date = normaliseWorkforceDate_(suggestion.Date);
    const personName = String(suggestion["Suggested Employee Name"] || "");
    const personKey = normaliseWorkforcePerson_(personName);
    const datePersonKey = date + "|" + personKey;
    if (usedByDate[datePersonKey]) return;
    const person = staffByName[personKey] || {};
    usedByDate[datePersonKey] = true;
    assignedGaps[gapId] = true;
    assignments.push({
      "Assignment ID": "relief_assignment_" + gapId + "_" + slugifyWorkforce_(personName),
      "Gap ID": gapId,
      "Site ID": gap["Site ID"],
      "Site Name": gap["Site Name"],
      "Date": date,
      "Weekday": gap.Weekday,
      "Role": gap.Role,
      "Covering Employee ID": person["Employee ID"] || suggestion["Suggested Employee ID"],
      "Covering Employee Name": personName,
      "Covering Email": person.Email || "",
      "Covered Employee Name": gap["Employee Name"],
      "Status": "Draft",
      "Score": suggestion.Score,
      "Reason": suggestion.Reason,
      "Generated At": now,
      "Notes": "Generated from relief rota engine"
    });
  });
  const sheet = ensureReliefAssignmentsSheet_(spreadsheet);
  replaceWorkforceSheetRowsFast_(sheet, assignments);
  return {
    ok: true,
    suggestionsCreated: suggestionResult.suggestionsCreated,
    assignmentsCreated: assignments.length,
    message: assignments.length + " draft relief assignment(s) generated."
  };
}

function generateReliefRotaPdfForWeek_(siteId, weekStartDate) {
  const spreadsheet = getWorkforceSpreadsheet_();
  const cleanSiteId = String(siteId || "").trim();
  if (!cleanSiteId) throw new Error("Choose a site first.");
  const weekStart = normaliseWeekStart_(weekStartDate);
  const weekDates = getWeekDates_(weekStart).filter(function(day) {
    return ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].indexOf(day.weekday) !== -1;
  });
  const dateSet = {};
  weekDates.forEach(function(day) { dateSet[day.date] = true; });

  const cockpit = buildRotaCockpit_(cleanSiteId, weekStart);
  const gapsById = {};
  (cockpit.days || []).forEach(function(day) {
    (day.gaps || []).forEach(function(gap) {
      if (gap.gapId || gap["Gap ID"]) gapsById[String(gap.gapId || gap["Gap ID"])] = gap;
    });
  });

  const suggestions = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.reliefSuggestions)
  ).filter(function(suggestion) {
    return String(suggestion["Site ID"] || "") === cleanSiteId &&
      dateSet[normaliseWorkforceDate_(suggestion.Date)] &&
      String(suggestion.Approved || "").toUpperCase() !== "TRUE";
  }).sort(function(a, b) {
    return String(a.Date || "").localeCompare(String(b.Date || "")) ||
      Number(b.Score || 0) - Number(a.Score || 0);
  });

  const staffByName = {};
  readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory))
    .forEach(function(person) {
      staffByName[normaliseWorkforcePerson_(person.Name)] = person;
    });

  const usedByDate = {};
  const assignedGaps = {};
  const now = new Date();
  const assignments = [];
  suggestions.forEach(function(suggestion) {
    const gapId = String(suggestion["Gap ID"] || "");
    const gap = gapsById[gapId] || {};
    if (!gapId || assignedGaps[gapId]) return;
    const date = normaliseWorkforceDate_(suggestion.Date);
    const personName = String(suggestion["Suggested Employee Name"] || "");
    const personKey = normaliseWorkforcePerson_(personName);
    const datePersonKey = date + "|" + personKey;
    if (usedByDate[datePersonKey]) return;
    const person = staffByName[personKey] || {};
    usedByDate[datePersonKey] = true;
    assignedGaps[gapId] = true;
    assignments.push({
      "Assignment ID": "relief_assignment_" + gapId + "_" + slugifyWorkforce_(personName),
      "Gap ID": gapId,
      "Site ID": cleanSiteId,
      "Site Name": cockpit.siteName || cleanSiteId,
      "Date": date,
      "Weekday": gap.weekday || gap.Weekday || getWorkforceWeekday_(new Date(date + "T00:00:00")),
      "Role": gap.role || gap.Role || suggestion.Role,
      "Covering Employee ID": person["Employee ID"] || suggestion["Suggested Employee ID"],
      "Covering Employee Name": personName,
      "Covering Email": person.Email || "",
      "Covered Employee Name": gap.missingEmployeeName || gap.employeeName || gap["Employee Name"] || "",
      "Status": "Draft",
      "Score": suggestion.Score,
      "Reason": suggestion.Reason,
      "Generated At": now,
      "Notes": "Generated from selected week rota cockpit"
    });
  });

  if (assignments.length) {
    upsertWorkforceRows_(ensureReliefAssignmentsSheet_(spreadsheet), "Assignment ID", assignments);
  }

  const filename = [
    "FIKA relief rota",
    cockpit.siteName || cleanSiteId,
    "WC " + weekStart
  ].join(" - ") + ".pdf";
  const pdf = buildReliefRotaPdfBlob_(cockpit, assignments, weekDates, filename);
  return {
    ok: true,
    siteId: cleanSiteId,
    siteName: cockpit.siteName || cleanSiteId,
    weekStart: weekStart,
    weekEnd: weekDates.length ? weekDates[weekDates.length - 1].date : weekStart,
    assignmentsCreated: assignments.length,
    filename: filename,
    mimeType: "application/pdf",
    base64: Utilities.base64Encode(pdf.getBytes()),
    message: assignments.length
      ? assignments.length + " relief assignment(s) generated for the selected week."
      : "No relief assignments were needed for the selected week."
  };
}

function buildReliefRotaPdfBlob_(cockpit, assignments, weekDates, filename) {
  const assignmentsByDate = {};
  (assignments || []).forEach(function(row) {
    const date = normaliseWorkforceDate_(row.Date);
    if (!assignmentsByDate[date]) assignmentsByDate[date] = [];
    assignmentsByDate[date].push(row);
  });
  const html = [
    "<html><head><style>",
    "body{font-family:Arial,sans-serif;color:#17113f;margin:28px}",
    "h1{font-size:24px;margin:0 0 4px} h2{font-size:15px;margin:0 0 18px;color:#746f91}",
    "table{width:100%;border-collapse:collapse;margin:12px 0 22px}",
    "th,td{border:1px solid #ddd8ef;padding:8px;text-align:left;font-size:11px;vertical-align:top}",
    "th{background:#241170;color:#fff} .empty{color:#746f91;font-style:italic}",
    ".day{margin-top:18px} .pill{display:inline-block;padding:4px 8px;border-radius:999px;background:#e9fff3;color:#277a55;font-size:11px;font-weight:bold}",
    "</style></head><body>",
    "<h1>FIKA Relief Rota</h1>",
    "<h2>" + escapeHtmlForPdf_(cockpit.siteName || cockpit.siteId || "Site") + " · Week commencing " + escapeHtmlForPdf_(cockpit.weekStart || "") + "</h2>",
    "<p class='pill'>" + String((assignments || []).length) + " draft assignment(s)</p>"
  ];
  (weekDates || []).forEach(function(day) {
    const rows = assignmentsByDate[day.date] || [];
    html.push("<div class='day'><h3>" + escapeHtmlForPdf_(day.weekday) + " " + escapeHtmlForPdf_(day.date) + "</h3>");
    if (!rows.length) {
      html.push("<p class='empty'>No relief cover generated for this day.</p></div>");
      return;
    }
    html.push("<table><thead><tr><th>Role</th><th>Covering</th><th>Covering email</th><th>Covering for</th><th>Reason</th></tr></thead><tbody>");
    rows.forEach(function(row) {
      html.push("<tr><td>" + escapeHtmlForPdf_(row.Role) + "</td><td>" +
        escapeHtmlForPdf_(row["Covering Employee Name"]) + "</td><td>" +
        escapeHtmlForPdf_(row["Covering Email"]) + "</td><td>" +
        escapeHtmlForPdf_(row["Covered Employee Name"]) + "</td><td>" +
        escapeHtmlForPdf_(row.Reason) + "</td></tr>");
    });
    html.push("</tbody></table></div>");
  });
  html.push("</body></html>");
  return Utilities.newBlob(html.join(""), "text/html", filename.replace(/\.pdf$/i, ".html"))
    .getAs("application/pdf")
    .setName(filename);
}

function escapeHtmlForPdf_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function createAgencyRequestForGap(gapId, agencyId) {
  const spreadsheet = getWorkforceSpreadsheet_();
  const gaps = readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.coverageGaps));
  const agencies = readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.agencyContacts));
  const gap = gaps.filter(function(row) { return String(row["Gap ID"]) === String(gapId); })[0];
  if (!gap) throw new Error("Gap was not found.");
  const agency = agencyId
    ? agencies.filter(function(row) { return String(row["Agency ID"]) === String(agencyId); })[0]
    : agencies.filter(function(row) { return workforceBoolean_(row.Active); })[0];
  if (!agency) throw new Error("Add at least one active agency contact first.");
  const requestId = "agency_request_" + gap["Gap ID"];
  const row = {
    "Agency Request ID": requestId,
    "Site ID": gap["Site ID"],
    "Date": normaliseWorkforceDate_(gap.Date),
    "Role": gap.Role,
    "Agency": agency["Agency Name"],
    "Rate": agency["Default Rate"],
    "Hours": "",
    "Status": "Draft",
    "Requested By": Session.getActiveUser().getEmail(),
    "Notes": "Created from gap " + gap["Gap ID"] + " | Contact: " + (agency.Email || "")
  };
  upsertWorkforceRows_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.agencyRequests), "Agency Request ID", [row]);
  return {
    ok: true,
    agencyRequestId: requestId,
    message: "Draft agency request created."
  };
}

function ensureReliefSuggestionsSheet_(spreadsheet) {
  return setupWorkforceSheet_(spreadsheet, WORKFORCE_CONFIG.sheets.reliefSuggestions, [
    "Suggestion ID", "Gap ID", "Site ID", "Date", "Role", "Suggested Employee ID",
    "Suggested Employee Name", "Reason", "Score", "Reviewed", "Approved"
  ]);
}

function ensureReliefAssignmentsSheet_(spreadsheet) {
  return setupWorkforceSheet_(spreadsheet, WORKFORCE_CONFIG.sheets.reliefAssignments, [
    "Assignment ID", "Gap ID", "Site ID", "Site Name", "Date", "Weekday",
    "Role", "Covering Employee ID", "Covering Employee Name", "Covering Email",
    "Covered Employee Name", "Status", "Score", "Reason", "Generated At", "Notes"
  ]);
}

function replaceWorkforceSheetRowsFast_(sheet, rowObjects) {
  if (!sheet) throw new Error("Target workforce sheet was not found.");
  clearWorkforceSheetData_(sheet);
  if (!rowObjects || !rowObjects.length) return;
  const map = workforceHeaderMap_(sheet);
  const headers = Object.keys(map);
  const values = rowObjects.map(function(rowObject) {
    return headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(rowObject, header)
        ? rowObject[header]
        : "";
    });
  });
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function buildCoverageGapIndex_(gaps) {
  const index = {};
  gaps.forEach(function(gap) {
    const date = normaliseWorkforceDate_(gap.Date);
    if (!date) return;
    const key = [
      String(gap["Site ID"] || ""),
      date,
      String(gap.Weekday || ""),
      normaliseWorkforcePerson_(gap["Employee Name"])
    ].join("|");
    index[key] = gap;
  });
  return index;
}

function scoreReliefCandidate_(person, gap, historySignal) {
  let score = 10;
  historySignal = historySignal || {};
  const role = String(gap.Role || "").toLowerCase();
  const personRole = String(person.Role || "").toLowerCase();
  if (personRole && role && (personRole.indexOf(role) !== -1 || role.indexOf(personRole) !== -1)) score += 50;
  if (workforceBoolean_(person["Relief Team"])) score += 25;
  if (String(person["Primary Site"] || "") === String(gap["Site Name"] || "")) score += 20;
  if (String(person["Secondary Sites"] || "").indexOf(String(gap["Site Name"] || "")) !== -1) score += 12;
  if (/chef/.test(role) && /chef/.test(personRole)) score += 25;
  if (/barista/.test(role) && /barista/.test(personRole)) score += 25;
  if (/hospitality/.test(role) && /hospitality/.test(personRole)) score += 20;
  score += Math.min(140, Number(historySignal.score || 0));
  return score;
}

function buildReliefReason_(person, gap, reliefMatch, historySignal) {
  const reasons = [];
  historySignal = historySignal || {};
  if (historySignal.exactCoverCount) {
    reasons.push("History: covered " + gap["Employee Name"] + " " + historySignal.exactCoverCount + " time(s)");
  } else if (historySignal.siteRoleCoverCount) {
    reasons.push("History: covered this site/role " + historySignal.siteRoleCoverCount + " time(s)");
  } else if (historySignal.roleCoverCount) {
    reasons.push("History: covered this role " + historySignal.roleCoverCount + " time(s)");
  }
  if (reliefMatch) reasons.push("Relief rota: " + reliefMatch["Site Name"] + " " + reliefMatch.Shift);
  if (person.Role) reasons.push("Role: " + person.Role);
  if (workforceBoolean_(person["Relief Team"])) reasons.push("Relief team");
  if (person["Primary Site"]) reasons.push("Primary site: " + person["Primary Site"]);
  return reasons.join(" | ") || "Available staff member";
}

function readCoverHistoryRows_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.coverHistory);
  return sheet ? readWorkforceObjects_(sheet) : [];
}

function buildCoverHistoryIndex_(rows) {
  const index = {};
  (rows || []).forEach(function(row) {
    if (String(row.Outcome || "").toLowerCase() === "cancelled") return;
    const covering = normaliseWorkforcePerson_(row["Covering Employee Name"]);
    const covered = normaliseWorkforcePerson_(row["Covered Employee Name"]);
    const siteId = String(row["Site ID"] || "");
    const role = String(row.Role || "").toLowerCase();
    if (!covering) return;
    incrementCoverHistoryIndex_(index, ["covering", covering].join("|"), row);
    if (covered) incrementCoverHistoryIndex_(index, ["covered", covered, "covering", covering].join("|"), row);
    if (covered && siteId) incrementCoverHistoryIndex_(index, ["covered", covered, "site", siteId, "covering", covering].join("|"), row);
    if (covered && role) incrementCoverHistoryIndex_(index, ["covered", covered, "role", role, "covering", covering].join("|"), row);
    if (siteId && role) incrementCoverHistoryIndex_(index, ["site", siteId, "role", role, "covering", covering].join("|"), row);
    if (role) incrementCoverHistoryIndex_(index, ["role", role, "covering", covering].join("|"), row);
  });
  return index;
}

function incrementCoverHistoryIndex_(index, key, row) {
  if (!index[key]) index[key] = { count: 0, lastDate: "" };
  index[key].count += 1;
  const date = normaliseWorkforceDate_(row.Date);
  if (date && (!index[key].lastDate || date > index[key].lastDate)) index[key].lastDate = date;
}

function getCoverHistorySignal_(index, gap, person) {
  index = index || {};
  const covering = normaliseWorkforcePerson_(person.Name || person["Employee Name"]);
  const covered = normaliseWorkforcePerson_(gap["Employee Name"]);
  const siteId = String(gap["Site ID"] || "");
  const role = String(gap.Role || "").toLowerCase();
  const exact = index[["covered", covered, "covering", covering].join("|")] || {};
  const exactSite = index[["covered", covered, "site", siteId, "covering", covering].join("|")] || {};
  const exactRole = index[["covered", covered, "role", role, "covering", covering].join("|")] || {};
  const siteRole = index[["site", siteId, "role", role, "covering", covering].join("|")] || {};
  const roleOnly = index[["role", role, "covering", covering].join("|")] || {};
  const exactCoverCount = Number(exact.count || 0);
  const siteRoleCoverCount = Number(siteRole.count || 0);
  const roleCoverCount = Number(roleOnly.count || 0);
  const score =
    Math.min(110, exactCoverCount * 45) +
    Math.min(35, Number(exactSite.count || 0) * 12) +
    Math.min(25, Number(exactRole.count || 0) * 8) +
    Math.min(30, siteRoleCoverCount * 10) +
    Math.min(15, roleCoverCount * 4);
  return {
    exactCoverCount: exactCoverCount,
    siteRoleCoverCount: siteRoleCoverCount,
    roleCoverCount: roleCoverCount,
    score: score,
    lastCoveredDate: exact.lastDate || exactSite.lastDate || exactRole.lastDate || siteRole.lastDate || roleOnly.lastDate || ""
  };
}

function buildReliefAvailabilityIndex_(rows) {
  const index = {};
  rows.forEach(function(row) {
    const date = normaliseWorkforceDate_(row.Date);
    if (!date) return;
    if (String(row.Status || "").toLowerCase() === "unavailable") return;
    if (String(row.Status || "").toLowerCase() === "bank holiday") return;
    if (!index[date]) index[date] = [];
    index[date].push(row);
  });
  return index;
}

function findReliefAvailabilityForPerson_(rows, personName) {
  const clean = normaliseWorkforcePerson_(personName);
  return (rows || []).filter(function(row) {
    return normaliseWorkforcePerson_(row["Relief Name"]) === clean;
  })[0] || null;
}

function buildRotaCoverIndex_(exceptions) {
  const index = {};
  exceptions.forEach(function(exception) {
    if (!isGapClearedByException_(exception)) return;
    const date = normaliseWorkforceDate_(exception.Date);
    if (!date) return;
    const role = String(exception.Role || "");
    if (!role) return;
    index[getCoverDateRoleKey_(exception["Site ID"], date, role)] = exception;
  });
  return index;
}

function getCoverDateRoleKey_(siteId, dateText, role) {
  return [
    String(siteId || ""),
    normaliseWorkforceDate_(dateText),
    normaliseRoleForCover_(role)
  ].join("|");
}

function normaliseRoleForCover_(role) {
  return String(role || "").toLowerCase().replace(/\s+/g, " ").trim();
}

function uniqueWorkforceRotaTemplates_(templates) {
  const byKey = {};
  (templates || []).forEach(function(template) {
    const key = getRotaTemplateDuplicateKey_(template);
    if (!key) return;
    const existing = byKey[key];
    if (!existing ||
        String(existing.Source || "") !== "Web App" ||
        String(template.Source || "") === "Web App") {
      byKey[key] = template;
    }
  });
  return Object.keys(byKey).map(function(key) { return byKey[key]; });
}

function getRotaTemplateDuplicateKey_(template) {
  return [
    String(template["Site ID"] || "").trim().toLowerCase(),
    String(template.Weekday || "").trim().toLowerCase(),
    normaliseRoleForCover_(template.Role),
    normaliseWorkforcePerson_(template["Employee Name"]),
    String(template["Standard Status"] || "").trim().toUpperCase()
  ].join("|");
}

function normaliseWeekStart_(value) {
  const dateText = normaliseWorkforceDate_(value) || getNextMonday_();
  const date = new Date(dateText + "T00:00:00");
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  const monday = new Date(date.getTime() + diff * 86400000);
  return Utilities.formatDate(monday, WORKFORCE_CONFIG.timeZone, "yyyy-MM-dd");
}

function getWeekDates_(weekStart) {
  const start = new Date(weekStart + "T00:00:00");
  const days = [];
  for (let index = 0; index < 7; index++) {
    const date = new Date(start.getTime() + index * 86400000);
    days.push({
      date: Utilities.formatDate(date, WORKFORCE_CONFIG.timeZone, "yyyy-MM-dd"),
      weekday: getWorkforceWeekday_(date)
    });
  }
  return days;
}

function getNextMonday_() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const day = today.getDay();
  const add = day === 1 ? 0 : (8 - day) % 7 || 7;
  return Utilities.formatDate(new Date(today.getTime() + add * 86400000), WORKFORCE_CONFIG.timeZone, "yyyy-MM-dd");
}

function clearWorkforceSheetData_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
}
