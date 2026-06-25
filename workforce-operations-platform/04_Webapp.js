function doGet() {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle(WORKFORCE_CONFIG.appName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWorkforceDashboardData() {
  try {
    const status = safeGetBrightHrApiStatus_();
    const properties = PropertiesService.getScriptProperties();
    const spreadsheet = safeGetWorkforceSpreadsheet_();
    const summary = spreadsheet ? getWorkforceSummary_(spreadsheet) : getEmptyWorkforceSummary_();
    const gaps = spreadsheet ? safeGetCoverageGapSummary_() : { count: 0, gaps: [] };
    const relief = spreadsheet ? safeGetReliefSuggestionSummary_() : { count: 0, suggestions: [] };
    const options = spreadsheet ? safeGetRotaAppOptions_() : { sites: [], defaultWeekStart: "" };
    const settings = spreadsheet ? safeGetWorkforceSettingsData_() : getEmptyWorkforceSettingsData_();
    return toPlainWorkforceResponse_(buildWorkforceDashboardResponse_({
      ok: true,
      brightHr: status,
      legacyRota: {
        configured: Boolean(
          properties.getProperty(WORKFORCE_CONFIG.scriptProperties.legacyRotaSpreadsheetId)
        )
      },
      reliefRota: {
        configured: Boolean(
          properties.getProperty(WORKFORCE_CONFIG.scriptProperties.reliefRotaSpreadsheetId)
        )
      },
      spreadsheet: spreadsheet ? {
        configured: true,
        name: spreadsheet.getName(),
        url: spreadsheet.getUrl()
      } : {
        configured: false,
        name: "",
        url: ""
      },
      summary: summary,
      gaps: gaps,
      relief: relief,
      options: options,
      settings: settings
    }));
  } catch (error) {
    return toPlainWorkforceResponse_(buildWorkforceDashboardResponse_({
      ok: false,
      error: error.message || String(error),
      brightHr: safeGetBrightHrApiStatus_(),
      spreadsheet: { configured: false, name: "", url: "" },
      summary: getEmptyWorkforceSummary_(),
      gaps: { count: 0, gaps: [] },
      relief: { count: 0, suggestions: [] },
      options: { sites: [], defaultWeekStart: "" },
      settings: getEmptyWorkforceSettingsData_()
    }));
  }
}

function getWorkforceDashboardSummaryFromUi() {
  try {
    const status = safeGetBrightHrApiStatus_();
    const properties = PropertiesService.getScriptProperties();
    const spreadsheet = safeGetWorkforceSpreadsheet_();
    const response = buildWorkforceDashboardResponse_({
      ok: true,
      brightHr: status,
      legacyRota: {
        configured: Boolean(
          properties.getProperty(WORKFORCE_CONFIG.scriptProperties.legacyRotaSpreadsheetId)
        )
      },
      reliefRota: {
        configured: Boolean(
          properties.getProperty(WORKFORCE_CONFIG.scriptProperties.reliefRotaSpreadsheetId)
        )
      },
      spreadsheet: spreadsheet ? getWorkforceSpreadsheetMeta_(spreadsheet) : {
        configured: false,
        name: "",
        url: ""
      },
      summary: spreadsheet ? getWorkforceSummary_(spreadsheet) : getEmptyWorkforceSummary_(),
      options: spreadsheet ? safeGetRotaAppOptions_() : { sites: [], defaultWeekStart: "" }
    });
    delete response.gaps;
    delete response.relief;
    delete response.settings;
    return toPlainJson_(response);
  } catch (error) {
    const response = buildWorkforceDashboardResponse_({
      ok: false,
      error: error.message || String(error),
      brightHr: safeGetBrightHrApiStatus_(),
      spreadsheet: { configured: false, name: "", url: "" },
      summary: getEmptyWorkforceSummary_(),
      options: { sites: [], defaultWeekStart: "" }
    });
    delete response.gaps;
    delete response.relief;
    delete response.settings;
    return toPlainJson_(response);
  }
}

function getCoverageGapsFromUi(limit) {
  return toPlainJson_(safeGetCoverageGapSummary_(limit || 40));
}

function getReliefSuggestionsForGapFromUi(gapId, limit) {
  try {
    const spreadsheet = getWorkforceSpreadsheet_();
    const cleanGapId = String(gapId || "").trim();
    if (!cleanGapId) return { ok: true, count: 0, suggestions: [] };
    const suggestions = readWorkforceObjects_(
      spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.reliefSuggestions)
    ).filter(function(row) {
      return String(row["Gap ID"] || "").trim() === cleanGapId &&
        String(row.Approved || "").toUpperCase() !== "TRUE";
    }).sort(function(a, b) {
      return Number(b.Score || 0) - Number(a.Score || 0);
    });
    return toPlainJson_({
      ok: true,
      gapId: cleanGapId,
      count: suggestions.length,
      suggestions: suggestions.slice(0, Math.max(1, Number(limit || 8)))
    });
  } catch (error) {
    return toPlainJson_({
      ok: false,
      error: error.message || String(error),
      gapId: String(gapId || ""),
      count: 0,
      suggestions: []
    });
  }
}

function getSettingsLibraryFromUi() {
  try {
    const spreadsheet = safeGetWorkforceSpreadsheet_();
    return toPlainJson_({
      ok: true,
      settings: spreadsheet ? safeGetWorkforceSettingsData_() : getEmptyWorkforceSettingsData_(),
      options: spreadsheet ? safeGetRotaAppOptions_() : { sites: [], defaultWeekStart: "" }
    });
  } catch (error) {
    return toPlainJson_({
      ok: false,
      error: error.message || String(error),
      settings: getEmptyWorkforceSettingsData_(),
      options: { sites: [], defaultWeekStart: "" }
    });
  }
}

function getStaffDirectoryLiteFromUi() {
  try {
    return toPlainJson_({
      ok: true,
      staff: getStaffDirectoryLite_()
    });
  } catch (error) {
    return toPlainJson_({
      ok: false,
      error: error.message || String(error),
      staff: []
    });
  }
}

function getRotaBoardFromUi(siteId, weekStartDate) {
  return toPlainJson_(getWeeklyRotaBoard(siteId, weekStartDate));
}

function getRotaCockpitFromUi(siteId, weekStartDate) {
  try {
    return toPlainJson_(buildRotaCockpit_(siteId, weekStartDate));
  } catch (error) {
    return toPlainJson_({
      ok: false,
      error: error.message || String(error),
      siteId: String(siteId || ""),
      weekStart: String(weekStartDate || ""),
      days: [],
      totals: getEmptyRotaCockpitTotals_()
    });
  }
}

function approveSuggestedCoverFromUi(gapId) {
  const suggestion = getBestReliefSuggestionForGap_(gapId);
  if (!suggestion || !suggestion.employeeName) {
    throw new Error("No suggested cover was found for this gap.");
  }
  return assignCoverForGap_(gapId, {
    coverName: suggestion.employeeName,
    coverType: "Relief",
    notes: [
      "Approved from rota cockpit",
      suggestion.reason,
      suggestion.score ? "Score " + suggestion.score : ""
    ].filter(Boolean).join(" | ")
  });
}

function chooseCoverForGapFromUi(gapId, employeeId) {
  const person = getStaffLiteByEmployeeId_(employeeId);
  if (!person || !person.name) throw new Error("Could not find that employee.");
  return assignCoverForGap_(gapId, {
    coverName: person.name,
    coverType: "Relief",
    notes: "Chosen manually from rota cockpit"
  });
}

function requestAgencyForGapFromUi(gapId, agencyId, details) {
  if (agencyId) return requestAgencyForGap_(gapId, agencyId, details || {});
  const defaultAgency = getDefaultAgencyContact_();
  if (defaultAgency && defaultAgency["Agency ID"]) {
    return requestAgencyForGap_(gapId, defaultAgency["Agency ID"], details || {});
  }
  return markCoverageGap_(gapId, "agency_requested", {
    notes: "Agency requested from rota cockpit. No default agency contact was configured."
  });
}

function markGapResolvedFromUi(gapId) {
  return markCoverageGap_(gapId, "resolved", {
    notes: "Resolved from rota cockpit"
  });
}

function snoozeGapFromUi(gapId) {
  return markCoverageGap_(gapId, "snoozed", {
    notes: "Snoozed from rota cockpit"
  });
}

function testBrightHrAbsenceEndpointFromUi() {
  try {
    const result = syncBrightHrAbsences();
    return toPlainJson_({
      ok: true,
      message: "Absence endpoint worked. " + (result.synced || 0) + " absence row(s) updated.",
      result: result
    });
  } catch (error) {
    return toPlainJson_({
      ok: false,
      error: error.message || String(error),
      message: "Absence endpoint test failed: " + (error.message || String(error))
    });
  }
}

function getWorkforceDashboardPing() {
  return {
      ok: true,
      appName: WORKFORCE_CONFIG.appName,
      appVersion: WORKFORCE_CONFIG.appVersion || "dev",
      timestamp: new Date().toISOString(),
      message: "Workforce web app server is responding."
  };
}

function setupWorkforceFromUi() {
  return setupWorkforceOperationsPlatform();
}

function syncBrightHrEmployeesFromUi() {
  return syncBrightHrEmployees();
}

function syncBrightHrAbsencesFromUi() {
  return syncBrightHrAbsences();
}

function discoverBrightHrAbsenceEndpointFromUi() {
  return discoverBrightHrAbsenceEndpoint();
}

function syncWorkforceFromUi() {
  const result = {
    ok: true,
    employees: null,
    absences: null,
    warnings: []
  };
  result.employees = syncBrightHrEmployees();
  try {
    result.absences = syncBrightHrAbsences();
  } catch (error) {
    result.warnings.push("Employee sync succeeded, but absence sync failed: " + (error.message || String(error)));
  }
  result.message = [
    "Workforce sync complete.",
    result.employees ? (result.employees.synced || 0) + " employee(s) updated." : "",
    result.absences ? (result.absences.synced || 0) + " absence row(s) updated." : "",
    result.warnings.length ? result.warnings.join(" ") : ""
  ].filter(Boolean).join(" ");
  return result;
}

function previewLegacyRotaImportFromUi() {
  return previewLegacyRotaCompressedImport(3);
}

function importLegacyRotaFromUi() {
  return importLegacyRotaCompressed();
}

function previewReliefRotaImportFromUi() {
  return previewReliefRotaImport();
}

function importReliefRotaFromUi() {
  return importReliefRota();
}

function clearLegacyRotaShiftNoiseFromUi() {
  const shiftResult = clearLegacyRotaShiftNoise();
  const compressedResult = clearLegacyCompressedImport();
  return {
    ok: true,
    shiftsDeleted: shiftResult.shiftsDeleted,
    templatesDeleted: compressedResult.templatesDeleted,
    exceptionsDeleted: compressedResult.exceptionsDeleted,
    agencyRequestsDeleted:
      (shiftResult.agencyRequestsDeleted || 0) +
      (compressedResult.agencyRequestsDeleted || 0),
    message: "Legacy import rows cleared from shifts, templates, exceptions and agency requests."
  };
}

function detectCoverageGapsFromUi() {
  return detectCoverageGaps(28);
}

function getWeeklyRotaBoardFromUi(siteId, weekStartDate) {
  return getRotaBoardFromUi(siteId, weekStartDate);
}

function generateReliefSuggestionsFromUi() {
  return generateReliefSuggestions(28);
}

function generateReliefRotaAssignmentsFromUi() {
  return generateReliefRotaAssignments(28);
}

function createAgencyRequestForGapFromUi(gapId, agencyId) {
  return createAgencyRequestForGap(gapId, agencyId);
}

function markCoverageGapFromUi(gapId, action, details) {
  return markCoverageGap_(gapId, action, details || {});
}

function assignCoverForGapFromUi(gapId, details) {
  return assignCoverForGap_(gapId, details || {});
}

function getWorkforceSettingsDataFromUi() {
  return getWorkforceSettingsData();
}

function saveWorkforceManagerFromUi(manager) {
  return saveWorkforceManager(manager);
}

function saveWorkforceAgencyContactFromUi(contact) {
  return saveWorkforceAgencyContact(contact);
}

function setReliefTeamMemberFromUi(employeeId, employeeName, isReliefTeam) {
  return setWorkforceReliefTeamMember(employeeId, employeeName, isReliefTeam);
}

function saveRotaTemplateFromUi(template) {
  return saveWorkforceRotaTemplate(template);
}

function dedupeRotaTemplatesFromUi() {
  return dedupeWorkforceRotaTemplates();
}

function safeGetWorkforceSpreadsheet_() {
  try {
    return getWorkforceSpreadsheet_();
  } catch (error) {
    return null;
  }
}

function safeGetBrightHrApiStatus_() {
  try {
    return getBrightHrApiStatus();
  } catch (error) {
    return {
      hasClientId: false,
      hasClientSecret: false,
      hasTokenUrl: false,
      hasApiBaseUrl: false,
      hasAbsencesPath: false,
      hasCachedAccessToken: false,
      tokenExpiresAt: ""
    };
  }
}

function safeGetCoverageGapSummary_(limit) {
  try {
    return getCoverageGapSummary(limit || 8);
  } catch (error) {
    return { count: 0, gaps: [], error: error.message || String(error) };
  }
}

function safeGetReliefSuggestionSummary_() {
  try {
    return getReliefSuggestionSummary(8);
  } catch (error) {
    return { count: 0, suggestions: [], error: error.message || String(error) };
  }
}

function safeGetRotaAppOptions_() {
  try {
    return getRotaAppOptions();
  } catch (error) {
    return { sites: [], defaultWeekStart: "", error: error.message || String(error) };
  }
}

function safeGetWorkforceSettingsData_() {
  try {
    return getWorkforceSettingsData();
  } catch (error) {
    const settings = getEmptyWorkforceSettingsData_();
    settings.error = error.message || String(error);
    return settings;
  }
}

function getEmptyWorkforceSettingsData_() {
  return {
    ok: false,
    managers: [],
    agencyContacts: [],
    roles: [],
    shiftPatterns: [],
    sites: [],
    staff: [],
    rotaTemplates: []
  };
}

function buildWorkforceDashboardResponse_(payload) {
  payload = payload || {};
  const summary = payload.summary || getEmptyWorkforceSummary_();
  return {
    ok: payload.ok !== false,
    error: payload.error || "",
    appName: WORKFORCE_CONFIG.appName,
    appVersion: WORKFORCE_CONFIG.appVersion || "dev",
    brightHr: payload.brightHr || safeGetBrightHrApiStatus_(),
    legacyRota: payload.legacyRota || { configured: false },
    reliefRota: payload.reliefRota || { configured: false },
    spreadsheet: payload.spreadsheet || { configured: false, name: "", url: "" },
    summary: summary,
    gaps: payload.gaps || { count: 0, gaps: [] },
    relief: payload.relief || { count: 0, suggestions: [] },
    options: payload.options || { sites: [], defaultWeekStart: "" },
    settings: payload.settings || getEmptyWorkforceSettingsData_(),
    modules: [
      { id: "staff", label: "Staff Directory", status: summary.staffCount ? "Live" : "Ready" },
      { id: "absence", label: "Absence Sync", status: summary.absenceCount ? "Live" : "Needs endpoint" },
      { id: "rota", label: "Rota Builder", status: "Next" },
      { id: "relief", label: "Relief Engine", status: "Next" },
      { id: "agency", label: "Agency Tracking", status: "Ready" }
    ]
  };
}

function toPlainWorkforceResponse_(response) {
  try {
    return JSON.parse(JSON.stringify(response || buildWorkforceDashboardResponse_({})));
  } catch (error) {
    return {
      ok: false,
      error: "Dashboard response could not be prepared: " + (error.message || String(error)),
      appName: WORKFORCE_CONFIG.appName,
      appVersion: WORKFORCE_CONFIG.appVersion || "dev",
      brightHr: {},
      legacyRota: { configured: false },
      reliefRota: { configured: false },
      spreadsheet: { configured: false, name: "", url: "" },
      summary: getEmptyWorkforceSummary_(),
      gaps: { count: 0, gaps: [] },
      relief: { count: 0, suggestions: [] },
      options: { sites: [], defaultWeekStart: "" },
      settings: getEmptyWorkforceSettingsData_(),
      modules: []
    };
  }
}

function toPlainJson_(value) {
  try {
    return JSON.parse(JSON.stringify(value || {}));
  } catch (error) {
    return {
      ok: false,
      error: "Response could not be prepared: " + (error.message || String(error))
    };
  }
}

function getWorkforceSpreadsheetMeta_(spreadsheet) {
  return {
    configured: true,
    name: spreadsheet.getName(),
    url: spreadsheet.getUrl()
  };
}

function getWorkforceSummary_(spreadsheet) {
  return {
    staffCount: getActiveStaffCount_(spreadsheet),
    absenceCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.absences),
    siteCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.sites),
    rotaShiftCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.rotaShifts),
    reliefSuggestionCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.reliefSuggestions),
    reliefAssignmentCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.reliefAssignments),
    agencyRequestCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.agencyRequests)
    ,
    managerCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.managers),
    agencyContactCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.agencyContacts),
    coverageGapCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.coverageGaps),
    rotaTemplateCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.rotaTemplates),
    rotaExceptionCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.rotaExceptions)
  };
}

function getEmptyWorkforceSummary_() {
  return {
    staffCount: 0,
    absenceCount: 0,
    siteCount: 0,
    rotaShiftCount: 0,
    reliefSuggestionCount: 0,
    reliefAssignmentCount: 0,
    agencyRequestCount: 0
    ,
    managerCount: 0,
    agencyContactCount: 0,
    coverageGapCount: 0,
    rotaTemplateCount: 0,
    rotaExceptionCount: 0
  };
}

function getWorkforceSheetRecordCount_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  return sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;
}

function getActiveStaffCount_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const map = workforceHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  const rowCount = lastRow - 1;
  const nameColumn = map.Name;
  if (!nameColumn) return 0;
  const names = sheet.getRange(2, nameColumn, rowCount, 1).getDisplayValues();
  const statuses = map["Employment Status"]
    ? sheet.getRange(2, map["Employment Status"], rowCount, 1).getDisplayValues()
    : [];
  const terminated = map.Terminated
    ? sheet.getRange(2, map.Terminated, rowCount, 1).getDisplayValues()
    : [];
  let count = 0;
  names.forEach(function(row, index) {
    const name = String(row[0] || "").trim();
    const employmentStatus = String((statuses[index] && statuses[index][0]) || "").toLowerCase();
    const isTerminated = workforceBoolean_((terminated[index] && terminated[index][0]) || "");
    if (name && employmentStatus !== "terminated" && !isTerminated) count += 1;
  });
  return count;
}

function getStaffDirectoryLite_() {
  const spreadsheet = getWorkforceSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const map = workforceHeaderMap_(sheet);
  const rowCount = sheet.getLastRow() - 1;
  const fields = [
    { key: "employeeId", header: "Employee ID" },
    { key: "name", header: "Name" },
    { key: "email", header: "Email" },
    { key: "role", header: "Role" },
    { key: "primarySite", header: "Primary Site" },
    { key: "secondarySites", header: "Secondary Sites" },
    { key: "contractHours", header: "Contract Hours" },
    { key: "employmentStatus", header: "Employment Status" },
    { key: "reliefTeam", header: "Relief Team" },
    { key: "eventTeam", header: "Event Team" },
    { key: "manager", header: "Manager" }
  ].filter(function(field) {
    return map[field.header];
  });
  if (!fields.length) return [];
  const columns = {};
  fields.forEach(function(field) {
    columns[field.key] = sheet.getRange(2, map[field.header], rowCount, 1).getDisplayValues();
  });
  const staff = [];
  for (let index = 0; index < rowCount; index += 1) {
    const person = {};
    fields.forEach(function(field) {
      person[field.key] = (columns[field.key][index] && columns[field.key][index][0]) || "";
    });
    if (!String(person.name || "").trim()) continue;
    person.reliefTeam = workforceBoolean_(person.reliefTeam);
    person.eventTeam = workforceBoolean_(person.eventTeam);
    person.manager = workforceBoolean_(person.manager);
    staff.push(person);
  }
  return staff;
}

function buildRotaCockpit_(siteId, weekStartDate) {
  const spreadsheet = getWorkforceSpreadsheet_();
  const cleanSiteId = String(siteId || "").trim();
  if (!cleanSiteId) throw new Error("Choose a site first.");
  const weekStart = normaliseWeekStart_(weekStartDate);
  const weekDates = getWeekDates_(weekStart);
  const weekEnd = weekDates[weekDates.length - 1].date;
  const site = getCockpitSite_(spreadsheet, cleanSiteId);
  const templates = uniqueWorkforceRotaTemplates_(readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.rotaTemplates)
  ).filter(function(template) {
    return String(template["Site ID"] || "") === cleanSiteId &&
      workforceBoolean_(template.Active);
  }));
  const absences = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.absences)
  );
  const exceptions = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.rotaExceptions)
  );
  const existingGaps = readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.coverageGaps)
  );
  const reliefAvailability = buildReliefAvailabilityIndex_(readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.reliefAvailability)
  ));
  const staff = getCockpitStaffCandidates_(spreadsheet, reliefAvailability);
  const absenceIndex = buildAbsenceIndex_(absences);
  const exceptionIndex = buildRotaExceptionIndex_(exceptions);
  const coverIndex = buildRotaCoverIndex_(exceptions);
  const existingGapIndex = buildExistingCockpitGapIndex_(existingGaps);
  const generatedGaps = [];
  const suggestions = [];
  const days = weekDates.map(function(day) {
    const dayTemplates = templates.filter(function(template) {
      return String(template.Weekday || "") === day.weekday &&
        String(template["Standard Status"] || "").toUpperCase() !== "OFF";
    });
    const scheduled = [];
    const dayAbsences = [];
    const dayGaps = [];
    dayTemplates.forEach(function(template) {
      const personName = String(template["Employee Name"] || "").trim();
      if (!personName) return;
      const templateKey = getGapTemplateDateKey_(template, day.date);
      const absence = absenceIndex[getGapPersonDateKey_(personName, day.date)];
      const exception = exceptionIndex[templateKey];
      const cover = coverIndex[getCoverDateRoleKey_(template["Site ID"], day.date, template.Role)];
      const isAbsent = Boolean(absence);
      const isCleared = Boolean(cover || (exception && isGapClearedByException_(exception)));
      const scheduleStatus = cover || (exception && isGapClearedByException_(exception))
        ? "covered"
        : isAbsent || exception ? "absent" : "scheduled";
      scheduled.push({
        employeeName: personName,
        role: String(template.Role || ""),
        status: scheduleStatus,
        coverName: cover ? String(cover["Employee Name"] || "") : "",
        coverType: cover ? String(cover["Exception Type"] || "Cover") : "",
        absenceType: absence ? String(absence["Absence Type"] || "Absence") : "",
        notes: exception ? String(exception.Notes || "") : ""
      });
      if (absence) {
        dayAbsences.push({
          employeeName: personName,
          role: String(template.Role || ""),
          absenceType: String(absence["Absence Type"] || "Absence"),
          startDate: normaliseWorkforceDate_(absence["Start Date"]),
          endDate: normaliseWorkforceDate_(absence["End Date"]) || normaliseWorkforceDate_(absence["Start Date"])
        });
      }
      if ((absence || (exception && String(exception["Exception Type"] || ""))) && !isCleared) {
        const gap = existingGapIndex[templateKey] || buildCoverageGap_(
          template,
          day.date,
          day.weekday,
          absence ? String(absence["Absence Type"] || "Absence") : String(exception["Exception Type"] || "Exception"),
          absence ? "BrightHR" : "Exception",
          exception ? exception.Notes || "" : ""
        );
        generatedGaps.push(gap);
        const gapSuggestions = getCockpitSuggestionsForGap_(gap, staff, absenceIndex, reliefAvailability);
        suggestions.push.apply(suggestions, gapSuggestions.rows);
        dayGaps.push({
          gapId: String(gap["Gap ID"] || ""),
          role: String(gap.Role || ""),
          missingEmployeeName: String(gap["Employee Name"] || ""),
          absenceType: String(gap["Gap Type"] || ""),
          priority: Number(gap.Priority || 3),
          status: String(gap.Status || "Open"),
          bestSuggestion: gapSuggestions.best
        });
      }
    });
    const openGaps = dayGaps.filter(function(gap) {
      return !/resolved|snoozed|ignored/i.test(String(gap.status || ""));
    });
    const dayStatus = !scheduled.length ? "closed"
      : openGaps.some(function(gap) { return !gap.bestSuggestion; }) ? "uncovered"
      : openGaps.length ? "cover_suggested"
      : "covered";
    return {
      date: day.date,
      weekday: day.weekday,
      status: dayStatus,
      scheduled: scheduled,
      absences: dayAbsences,
      gaps: dayGaps
    };
  });
  if (generatedGaps.length) {
    upsertWorkforceRows_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.coverageGaps), "Gap ID", generatedGaps);
  }
  if (suggestions.length) {
    upsertWorkforceRows_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.reliefSuggestions), "Suggestion ID", suggestions);
  }
  return {
    ok: true,
    siteId: cleanSiteId,
    siteName: site.siteName || cleanSiteId,
    weekStart: weekStart,
    weekEnd: weekEnd,
    days: days,
    totals: getRotaCockpitTotals_(days)
  };
}

function getCockpitSuggestionsForGap_(gap, staff, absenceIndex, reliefAvailability) {
  const date = normaliseWorkforceDate_(gap.Date);
  const availableRelief = reliefAvailability[date] || [];
  const candidates = staff.filter(function(person) {
    return String(person.Name || "") &&
      normaliseWorkforcePerson_(person.Name) !== normaliseWorkforcePerson_(gap["Employee Name"]) &&
      !absenceIndex[getGapPersonDateKey_(person.Name, date)];
  }).map(function(person) {
    const reliefMatch = findReliefAvailabilityForPerson_(availableRelief, person.Name);
    const score = scoreReliefCandidate_(person, gap) + (reliefMatch ? 60 : 0);
    return {
      person: person,
      reliefMatch: reliefMatch,
      score: score,
      reason: buildReliefReason_(person, gap, reliefMatch)
    };
  }).sort(function(a, b) {
    return b.score - a.score;
  }).slice(0, 3);
  const rows = candidates.map(function(candidate, index) {
    return {
      "Suggestion ID": ["relief", gap["Gap ID"], slugifyWorkforce_(candidate.person.Name), index + 1].join("_"),
      "Gap ID": gap["Gap ID"],
      "Site ID": gap["Site ID"],
      "Date": date,
      "Role": gap.Role,
      "Suggested Employee ID": candidate.person["Employee ID"],
      "Suggested Employee Name": candidate.person.Name,
      "Reason": candidate.reason,
      "Score": candidate.score,
      "Reviewed": false,
      "Approved": false
    };
  });
  const best = rows[0] ? {
    employeeId: String(rows[0]["Suggested Employee ID"] || ""),
    employeeName: String(rows[0]["Suggested Employee Name"] || ""),
    score: Number(rows[0].Score || 0),
    reason: String(rows[0].Reason || ""),
    confidence: getSuggestionConfidence_(Number(rows[0].Score || 0))
  } : null;
  return { best: best, rows: rows };
}

function getCockpitStaffCandidates_(spreadsheet, reliefAvailability) {
  const sheet = spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const map = workforceHeaderMap_(sheet);
  const rowCount = sheet.getLastRow() - 1;
  const headers = [
    "Employee ID",
    "Name",
    "Email",
    "Role",
    "Primary Site",
    "Secondary Sites",
    "Employment Status",
    "Terminated",
    "Relief Team",
    "Event Team"
  ].filter(function(header) {
    return map[header];
  });
  const columns = {};
  headers.forEach(function(header) {
    columns[header] = sheet.getRange(2, map[header], rowCount, 1).getDisplayValues();
  });
  const staff = [];
  for (let index = 0; index < rowCount; index += 1) {
    const person = {};
    headers.forEach(function(header) {
      person[header] = (columns[header][index] && columns[header][index][0]) || "";
    });
    if (String(person.Name || "").trim() && isReliefCandidate_(person, reliefAvailability)) {
      staff.push(person);
    }
  }
  return staff;
}

function getSuggestionConfidence_(score) {
  if (score >= 90) return "High";
  if (score >= 55) return "Medium";
  return "Low";
}

function getRotaCockpitTotals_(days) {
  const totals = getEmptyRotaCockpitTotals_();
  (days || []).forEach(function(day) {
    totals.scheduled += (day.scheduled || []).length;
    totals.absences += (day.absences || []).length;
    (day.gaps || []).forEach(function(gap) {
      if (/resolved|snoozed|ignored/i.test(String(gap.status || ""))) {
        totals.resolved += 1;
      } else if (gap.bestSuggestion) {
        totals.coverSuggested += 1;
      } else {
        totals.uncoveredGaps += 1;
      }
    });
  });
  return totals;
}

function getEmptyRotaCockpitTotals_() {
  return {
    scheduled: 0,
    absences: 0,
    uncoveredGaps: 0,
    coverSuggested: 0,
    resolved: 0
  };
}

function buildExistingCockpitGapIndex_(gaps) {
  const index = {};
  (gaps || []).forEach(function(gap) {
    const date = normaliseWorkforceDate_(gap.Date);
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

function getCockpitSite_(spreadsheet, siteId) {
  const site = readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.sites))
    .filter(function(row) {
      return String(row["Site ID"] || "") === String(siteId || "");
    })[0] || {};
  return {
    siteId: String(site["Site ID"] || siteId || ""),
    siteName: String(site["Site Name"] || siteId || "")
  };
}

function getBestReliefSuggestionForGap_(gapId) {
  const cleanGapId = String(gapId || "").trim();
  if (!cleanGapId) return null;
  const spreadsheet = getWorkforceSpreadsheet_();
  const suggestion = readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.reliefSuggestions))
    .filter(function(row) {
      return String(row["Gap ID"] || "").trim() === cleanGapId &&
        String(row.Approved || "").toUpperCase() !== "TRUE";
    })
    .sort(function(a, b) {
      return Number(b.Score || 0) - Number(a.Score || 0);
    })[0];
  return suggestion ? {
    employeeId: String(suggestion["Suggested Employee ID"] || ""),
    employeeName: String(suggestion["Suggested Employee Name"] || ""),
    score: Number(suggestion.Score || 0),
    reason: String(suggestion.Reason || "")
  } : null;
}

function getStaffLiteByEmployeeId_(employeeId) {
  const cleanId = String(employeeId || "").trim();
  if (!cleanId) return null;
  return getStaffDirectoryLite_().filter(function(person) {
    return String(person.employeeId || "").trim() === cleanId;
  })[0] || null;
}

function getDefaultAgencyContact_() {
  const spreadsheet = getWorkforceSpreadsheet_();
  return readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.agencyContacts))
    .filter(function(agency) {
      return agency["Agency ID"] && String(agency.Active || "").toUpperCase() !== "FALSE";
    })[0] || null;
}
