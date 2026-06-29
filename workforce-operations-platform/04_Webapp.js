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

function getRotaAssistantWeekFromUi(siteId, weekStartDate) {
  try {
    return toPlainJson_(buildRotaAssistantWeek_(siteId, weekStartDate));
  } catch (error) {
    return toPlainJson_({
      ok: false,
      error: error.message || String(error),
      siteId: String(siteId || ""),
      weekStart: String(weekStartDate || ""),
      generatedAt: new Date(),
      days: [],
      totals: getEmptyRotaAssistantTotals_()
    });
  }
}

function getPersonStandardRotaFromUi(employeeId) {
  return toPlainJson_({
    ok: true,
    employeeId: String(employeeId || ""),
    rows: getPersonStandardRota_(employeeId)
  });
}

function savePersonStandardRotaFromUi(employeeId, rows) {
  const result = savePersonStandardRota_(employeeId, rows || []);
  return toPlainJson_(result);
}

function getPersonRotaProfileFromUi(employeeId) {
  return toPlainJson_(getPersonRotaProfile_(employeeId));
}

function markEmployeeSickFromUi(employeeId, dateText, notes) {
  return toPlainJson_(markEmployeeSick_(employeeId, dateText, notes || ""));
}

function getPeopleStandardRotaForWeekFromUi(siteId, weekStartDate) {
  return toPlainJson_({
    ok: true,
    siteId: String(siteId || ""),
    weekStart: normaliseWeekStart_(weekStartDate),
    rows: getPeopleStandardRotaForWeek_(siteId, weekStartDate)
  });
}

function approveSuggestedCoverFromUi(gapId) {
  const suggestion = getBestReliefSuggestionForGap_(gapId);
  if (!suggestion || !suggestion.employeeName) {
    throw new Error("No suggested cover was found for this gap.");
  }
  return assignCoverForGap_(gapId, {
    coverEmployeeId: suggestion.employeeId,
    coverName: suggestion.employeeName,
    coverType: "Relief",
    notes: [
      "Approved from rota cockpit",
      suggestion.reason,
      suggestion.score ? "Score " + suggestion.score : ""
    ].filter(Boolean).join(" | ")
  });
}

function approveBestCoverFromUi(issueId) {
  return approveSuggestedCoverFromUi(issueId);
}

function chooseCoverForGapFromUi(gapId, employeeId) {
  const person = getStaffLiteByEmployeeId_(employeeId);
  if (!person || !person.name) throw new Error("Could not find that employee.");
  return assignCoverForGap_(gapId, {
    coverEmployeeId: person.employeeId,
    coverName: person.name,
    coverType: "Relief",
    notes: "Chosen manually from rota cockpit"
  });
}

function chooseCoverForIssueFromUi(issueId, employeeId) {
  return chooseCoverForGapFromUi(issueId, employeeId);
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

function requestAgencyForIssueFromUi(issueId) {
  return requestAgencyForGapFromUi(issueId);
}

function markGapResolvedFromUi(gapId) {
  return markCoverageGap_(gapId, "resolved", {
    notes: "Resolved from rota cockpit"
  });
}

function markIssueResolvedFromUi(issueId) {
  return markGapResolvedFromUi(issueId);
}

function ignoreIssueFromUi(issueId) {
  return markCoverageGap_(issueId, "ignored", {
    notes: "Ignored from Rota Assistant"
  });
}

function snoozeGapFromUi(gapId) {
  return markCoverageGap_(gapId, "snoozed", {
    notes: "Snoozed from rota cockpit"
  });
}

function snoozeIssueFromUi(issueId, untilDate) {
  return markCoverageGap_(issueId, "snoozed", {
    notes: "Snoozed from Rota Assistant" + (untilDate ? " until " + untilDate : "")
  });
}

function testBrightHrAbsenceEndpointFromUi() {
  return toPlainJson_(testBrightHrEndpointPreview_("absence"));
}

function testBrightHrConnectionFromUi() {
  try {
    const result = testBrightHrConnection();
    return toPlainJson_(result);
  } catch (error) {
    return toPlainJson_({
      ok: false,
      error: error.message || String(error),
      message: "BrightHR connection failed: " + (error.message || String(error))
    });
  }
}

function testBrightHrEmployeeEndpointFromUi() {
  return toPlainJson_(testBrightHrEndpointPreview_("employee"));
}

function getBrightHrAdminSettingsFromUi() {
  return toPlainJson_(getBrightHrAdminSettings_());
}

function saveBrightHrAdminSettingsFromUi(settings) {
  return toPlainJson_(saveBrightHrAdminSettings_(settings || {}));
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

function syncBrightHrAbsencesBatchFromUi(batchSize) {
  return toPlainJson_(syncBrightHrAbsencesBatch(batchSize));
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

function generateReliefRotaPdfForSelectedWeekFromUi(siteId, weekStartDate) {
  return toPlainJson_(generateReliefRotaPdfForWeek_(siteId, weekStartDate));
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
    rotaTemplates: [],
    personStandardRota: []
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
    coverHistoryCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.coverHistory),
    agencyRequestCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.agencyRequests),
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
    coverHistoryCount: 0,
    agencyRequestCount: 0,
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
  const weekDates = getWeekDates_(weekStart).filter(function(day) {
    return ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].indexOf(day.weekday) !== -1;
  });
  const weekEnd = weekDates[weekDates.length - 1].date;
  const site = getCockpitSite_(spreadsheet, cleanSiteId);
  const templates = uniqueWorkforceRotaTemplates_(readWorkforceObjects_(
    spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.rotaTemplates)
  ).filter(function(template) {
    return String(template["Site ID"] || "") === cleanSiteId &&
      workforceBoolean_(template.Active);
  }));
  const personStandardRows = getPeopleStandardRotaForWeek_(cleanSiteId, weekStart);
  const personTemplates = personStandardRows
    .filter(function(row) {
      return String(row.Site || "") === cleanSiteId;
    })
    .map(function(row) {
      return {
        "Template ID": ["person", row["Employee ID"], row["Day of Week"], cleanSiteId, row.Role].join("_"),
        "Site ID": cleanSiteId,
        "Site Name": site.siteName || cleanSiteId,
        "Weekday": row["Day of Week"],
        "Role": row.Role,
        "Employee ID": row["Employee ID"],
        "Employee Name": row["Employee Name"],
        "Start Time": row["Start Time"],
        "End Time": row["End Time"],
        "Standard Status": "IN",
        "Source": "Person Standard Rota",
        "Observations": row.Notes || "",
        "Active": true
      };
    });
  const combinedTemplates = uniqueCockpitTemplates_(templates.concat(personTemplates));
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
  const coverHistory = buildCoverHistoryIndex_(readCoverHistoryRows_(spreadsheet));
  const staff = getCockpitStaffCandidates_(spreadsheet, reliefAvailability);
  const staffLookup = buildCockpitStaffLiteLookup_(spreadsheet);
  const absenceIndex = buildAbsenceIndex_(absences);
  const exceptionIndex = buildRotaExceptionIndex_(exceptions);
  const coverIndex = buildRotaCoverIndex_(exceptions);
  const existingGapIndex = buildExistingCockpitGapIndex_(existingGaps);
  const generatedGaps = [];
  const suggestions = [];
  const days = weekDates.map(function(day) {
    const dayTemplates = combinedTemplates.filter(function(template) {
      return String(template.Weekday || "") === day.weekday &&
        String(template["Standard Status"] || "").toUpperCase() !== "OFF";
    });
    const fullRota = [];
    const dayAbsences = [];
    const dayGaps = [];
    const issues = [];
    dayTemplates.forEach(function(template) {
      const personName = String(template["Employee Name"] || "").trim();
      if (!personName) return;
      const employeeId = String(template["Employee ID"] || "") ||
        ((staffLookup.byName[normaliseWorkforcePerson_(personName)] || {}).employeeId || "");
      const templateKey = getGapTemplateDateKey_(template, day.date);
      const absence = absenceIndex[getGapPersonDateKey_(personName, day.date)];
      const exception = exceptionIndex[templateKey];
      const cover = coverIndex[getCoverDateRoleKey_(template["Site ID"], day.date, template.Role)];
      const isAbsent = Boolean(absence);
      const isCleared = Boolean(cover || (exception && isGapClearedByException_(exception)));
      const scheduleStatus = cover || (exception && isGapClearedByException_(exception))
        ? "covered"
        : isAbsent || exception ? "absent" : "scheduled";
      fullRota.push({
        employeeId: employeeId,
        employeeName: personName,
        role: String(template.Role || ""),
        siteId: String(template["Site ID"] || ""),
        siteName: String(template["Site Name"] || ""),
        startTime: String(template["Start Time"] || ""),
        endTime: String(template["End Time"] || ""),
        status: scheduleStatus,
        coverName: cover ? String(cover["Employee Name"] || "") : "",
        coverType: cover ? String(cover["Exception Type"] || "Cover") : "",
        absenceType: absence ? String(absence["Absence Type"] || "Absence") : "",
        notes: exception ? String(exception.Notes || "") : ""
      });
      if (absence) {
        const absenceIssue = {
          issueId: ["absence", cleanSiteId, day.date, slugifyWorkforce_(personName)].join("_"),
          type: "absence",
          employeeId: employeeId,
          employeeName: personName,
          role: String(template.Role || ""),
          reason: String(absence["Absence Type"] || "Absence"),
          absenceType: String(absence["Absence Type"] || "Absence"),
          startDate: normaliseWorkforceDate_(absence["Start Date"]),
          endDate: normaliseWorkforceDate_(absence["End Date"]) || normaliseWorkforceDate_(absence["Start Date"]),
          status: isCleared ? "resolved" : "open",
          coverName: cover ? String(cover["Employee Name"] || "") : "",
          primaryAction: isCleared ? "View cover" : "Find cover"
        };
        dayAbsences.push(absenceIssue);
        issues.push(absenceIssue);
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
        const gapSuggestions = getCockpitSuggestionsForGap_(gap, staff, absenceIndex, reliefAvailability, coverHistory);
        suggestions.push.apply(suggestions, gapSuggestions.rows);
        const gapIssue = {
          issueId: String(gap["Gap ID"] || ""),
          type: gapSuggestions.best ? "cover_suggested" : "uncovered",
          gapId: String(gap["Gap ID"] || ""),
          role: String(gap.Role || ""),
          employeeId: employeeId,
          missingEmployeeName: String(gap["Employee Name"] || ""),
          employeeName: String(gap["Employee Name"] || ""),
          absenceType: String(gap["Gap Type"] || ""),
          reason: String(gap["Gap Type"] || ""),
          priority: Number(gap.Priority || 3),
          status: String(gap.Status || "Open"),
          bestSuggestion: gapSuggestions.best,
          primaryAction: gapSuggestions.best ? "Approve cover" : "Request agency"
        };
        dayGaps.push(gapIssue);
        issues.push(gapIssue);
      }
      const conflict = getPersonStandardRotaConflict_(template, personStandardRows, cleanSiteId, day.weekday);
      if (conflict) issues.push(conflict);
    });
    const openGaps = dayGaps.filter(function(gap) {
      return !/resolved|snoozed|ignored/i.test(String(gap.status || ""));
    });
    const dayStatus = !fullRota.length ? "closed"
      : openGaps.some(function(gap) { return !gap.bestSuggestion; }) ? "uncovered"
      : openGaps.length ? "cover_suggested"
      : issues.some(function(issue) { return String(issue.type || "") === "conflict"; }) ? "cover_suggested"
      : "covered";
    return {
      date: day.date,
      weekday: day.weekday,
      status: dayStatus,
      summary: buildCockpitDaySummary_(dayStatus, issues, fullRota),
      issues: issues,
      fullRota: fullRota,
      scheduled: fullRota,
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

function buildRotaAssistantWeek_(siteId, weekStartDate) {
  const cockpit = buildRotaCockpit_(siteId, weekStartDate);
  const days = (cockpit.days || []).filter(function(day) {
    return ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
      .indexOf(String(day.weekday || "")) !== -1;
  }).map(function(day) {
    const issues = (day.issues || [])
      .filter(function(issue) {
        return !/resolved|snoozed|ignored|agency_requested/i.test(String(issue.status || ""));
      })
      .map(function(issue) {
        return formatAssistantIssue_(issue, day, cockpit);
      });
    const fullRota = (day.fullRota || []).map(function(row) {
      const issueIds = issues
        .filter(function(issue) {
          return normaliseWorkforcePerson_(issue.personName) === normaliseWorkforcePerson_(row.employeeName) ||
            String(issue.role || "") === String(row.role || "");
        })
        .map(function(issue) { return issue.issueId; });
      return {
        employeeId: String(row.employeeId || ""),
        employeeName: String(row.employeeName || ""),
        role: String(row.role || ""),
        startTime: String(row.startTime || ""),
        endTime: String(row.endTime || ""),
        source: String(row.source || row.Source || "site_standard"),
        hasIssue: Boolean(issueIds.length),
        issueIds: issueIds
      };
    });
    const summary = {
      scheduledCount: fullRota.length,
      absenceCount: issues.filter(function(issue) { return issue.issueType === "absence"; }).length,
      issueCount: issues.length,
      uncoveredCount: issues.filter(function(issue) { return issue.status === "needs_action"; }).length,
      suggestionCount: issues.filter(function(issue) { return issue.status === "suggested"; }).length
    };
    const status = !fullRota.length ? "closed"
      : summary.uncoveredCount ? "uncovered"
      : summary.suggestionCount ? "cover_suggested"
      : "covered";
    return {
      date: day.date,
      weekday: day.weekday,
      status: status,
      summary: summary,
      issues: issues,
      fullRota: fullRota
    };
  });
  return {
    ok: true,
    siteId: cockpit.siteId,
    siteName: cockpit.siteName,
    weekStart: cockpit.weekStart,
    weekEnd: cockpit.weekEnd,
    generatedAt: new Date(),
    totals: getRotaAssistantTotals_(days),
    days: days
  };
}

function formatAssistantIssue_(issue, day, cockpit) {
  const suggestion = issue.bestSuggestion ? {
    employeeId: String(issue.bestSuggestion.employeeId || ""),
    employeeName: String(issue.bestSuggestion.employeeName || ""),
    role: String(issue.bestSuggestion.role || ""),
    homeSite: String(issue.bestSuggestion.homeSite || ""),
    score: Number(issue.bestSuggestion.score || 0),
    confidence: String(issue.bestSuggestion.confidence || ""),
    reason: String(issue.bestSuggestion.reason || ""),
    matchBullets: getAssistantMatchBullets_(issue.bestSuggestion)
  } : null;
  const hasSuggestion = Boolean(suggestion && suggestion.employeeName);
  const rawType = String(issue.type || "");
  const issueType = rawType === "conflict" ? "conflict"
    : rawType === "absence" ? "absence"
    : hasSuggestion ? "suggestion"
    : "gap";
  return {
    issueId: String(issue.gapId || issue.issueId || ""),
    gapId: String(issue.gapId || issue.issueId || ""),
    issueType: issueType,
    type: hasSuggestion ? "cover_suggested" : rawType || "uncovered",
    status: hasSuggestion ? "suggested" : "needs_action",
    personId: String(issue.employeeId || ""),
    personName: String(issue.employeeName || issue.missingEmployeeName || ""),
    employeeId: String(issue.employeeId || ""),
    employeeName: String(issue.employeeName || issue.missingEmployeeName || ""),
    missingEmployeeName: String(issue.missingEmployeeName || issue.employeeName || ""),
    role: String(issue.role || ""),
    siteName: String(cockpit.siteName || cockpit.siteId || ""),
    date: day.date,
    weekday: day.weekday,
    absenceType: String(issue.absenceType || ""),
    absenceReason: String(issue.reason || issue.absenceType || ""),
    gapReason: String(issue.reason || issue.absenceType || "Needs cover"),
    reason: String(issue.reason || issue.absenceType || "Needs cover"),
    priority: Number(issue.priority || 3),
    bestSuggestion: suggestion
  };
}

function getAssistantMatchBullets_(suggestion) {
  const reason = String((suggestion && suggestion.reason) || "");
  return reason.split("|")
    .map(function(part) { return part.trim(); })
    .filter(Boolean)
    .slice(0, 4);
}

function getRotaAssistantTotals_(days) {
  const totals = getEmptyRotaAssistantTotals_();
  (days || []).forEach(function(day) {
    const summary = day.summary || {};
    totals.absenceCount += Number(summary.absenceCount || 0);
    totals.issueCount += Number(summary.issueCount || 0);
    totals.uncoveredGapCount += Number(summary.uncoveredCount || 0);
    totals.suggestionCount += Number(summary.suggestionCount || 0);
    totals.resolvedCount += Number(summary.resolvedCount || 0);
  });
  return totals;
}

function getEmptyRotaAssistantTotals_() {
  return {
    absenceCount: 0,
    issueCount: 0,
    uncoveredGapCount: 0,
    suggestionCount: 0,
    resolvedCount: 0
  };
}

function getCockpitSuggestionsForGap_(gap, staff, absenceIndex, reliefAvailability, coverHistory) {
  const date = normaliseWorkforceDate_(gap.Date);
  const availableRelief = reliefAvailability[date] || [];
  const candidates = staff.filter(function(person) {
    return String(person.Name || "") &&
      normaliseWorkforcePerson_(person.Name) !== normaliseWorkforcePerson_(gap["Employee Name"]) &&
      !absenceIndex[getGapPersonDateKey_(person.Name, date)];
  }).map(function(person) {
    const reliefMatch = findReliefAvailabilityForPerson_(availableRelief, person.Name);
    const historySignal = getCoverHistorySignal_(coverHistory, gap, person);
    const score = scoreReliefCandidate_(person, gap, historySignal) + (reliefMatch ? 60 : 0);
    return {
      person: person,
      reliefMatch: reliefMatch,
      historySignal: historySignal,
      score: score,
      reason: buildReliefReason_(person, gap, reliefMatch, historySignal)
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

function uniqueCockpitTemplates_(rows) {
  const seen = {};
  const output = [];
  (rows || []).forEach(function(row) {
    const key = [
      String(row["Site ID"] || ""),
      String(row.Weekday || ""),
      normaliseWorkforcePerson_(row["Employee Name"]),
      String(row.Role || "").toLowerCase()
    ].join("|");
    if (seen[key]) return;
    seen[key] = true;
    output.push(row);
  });
  return output;
}

function getPersonStandardRotaConflict_(template, personRows, selectedSiteId, weekday) {
  const employeeId = String(template["Employee ID"] || "").trim();
  const personName = normaliseWorkforcePerson_(template["Employee Name"]);
  const conflict = (personRows || []).filter(function(row) {
    const rowEmployeeId = String(row["Employee ID"] || "").trim();
    const samePerson = employeeId
      ? rowEmployeeId && rowEmployeeId === employeeId
      : normaliseWorkforcePerson_(row["Employee Name"]) === personName;
    return samePerson &&
      String(row["Day of Week"] || "") === String(weekday || "") &&
      String(row.Site || "") &&
      String(row.Site || "") !== String(selectedSiteId || "");
  })[0];
  if (!conflict) return null;
  return {
    issueId: ["conflict", selectedSiteId, weekday, slugifyWorkforce_(template["Employee Name"])].join("_"),
    type: "conflict",
    employeeId: employeeId || String(conflict["Employee ID"] || ""),
    employeeName: String(template["Employee Name"] || conflict["Employee Name"] || ""),
    role: String(template.Role || conflict.Role || ""),
    reason: "Person standard rota says " + String(conflict.Site || "") + " on " + String(weekday || ""),
    status: "open",
    primaryAction: "Edit person standard rota"
  };
}

function buildCockpitDaySummary_(status, issues, fullRota) {
  const issueRows = issues || [];
  const openIssues = issueRows.filter(function(issue) {
    return !/resolved|snoozed|ignored/i.test(String(issue.status || ""));
  });
  return {
    scheduledCount: (fullRota || []).length,
    absenceCount: issueRows.filter(function(issue) { return String(issue.type || "") === "absence"; }).length,
    issueCount: openIssues.length,
    uncoveredCount: openIssues.filter(function(issue) {
      return String(issue.type || "") === "uncovered";
    }).length,
    suggestionCount: openIssues.filter(function(issue) {
      return Boolean(issue.bestSuggestion);
    }).length,
    conflictCount: openIssues.filter(function(issue) { return String(issue.type || "") === "conflict"; }).length,
    resolvedCount: issueRows.length - openIssues.length,
    status: status || ""
  };
}

function getPersonStandardRota_(employeeId) {
  const cleanId = String(employeeId || "").trim();
  const spreadsheet = getWorkforceSpreadsheet_();
  return readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.personStandardRota))
    .filter(function(row) {
      return String(row["Employee ID"] || "").trim() === cleanId;
    });
}

function savePersonStandardRota_(employeeId, rows) {
  const cleanId = String(employeeId || "").trim();
  if (!cleanId) throw new Error("Employee ID is required.");
  const spreadsheet = getWorkforceSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.personStandardRota);
  const existing = readWorkforceObjects_(sheet).filter(function(row) {
    return String(row["Employee ID"] || "").trim() !== cleanId;
  });
  const nextRows = (rows || []).map(function(row) {
    return {
      "Employee ID": cleanId,
      "Employee Name": String(row["Employee Name"] || row.employeeName || "").trim(),
      "Day of Week": String(row["Day of Week"] || row.dayOfWeek || "").trim(),
      "Site": String(row.Site || row.site || "").trim(),
      "Role": String(row.Role || row.role || "").trim(),
      "Start Time": String(row["Start Time"] || row.startTime || "").trim(),
      "End Time": String(row["End Time"] || row.endTime || "").trim(),
      "Effective From": String(row["Effective From"] || row.effectiveFrom || "").trim(),
      "Effective To": String(row["Effective To"] || row.effectiveTo || "").trim(),
      "Active": row.Active === false ? false : String(row.Active || row.active || "TRUE").toUpperCase() !== "FALSE",
      "Notes": String(row.Notes || row.notes || "").trim()
    };
  }).filter(function(row) {
    return row["Employee Name"] && row["Day of Week"] && row.Site;
  });
  clearWorkforceSheetData_(sheet);
  appendWorkforceRows_(sheet, existing.concat(nextRows));
  return {
    ok: true,
    employeeId: cleanId,
    saved: nextRows.length,
    message: "Person standard rota saved."
  };
}

function appendWorkforceRows_(sheet, rows) {
  if (!sheet || !rows || !rows.length) return;
  const map = workforceHeaderMap_(sheet);
  const headers = Object.keys(map);
  const values = rows.map(function(row) {
    return headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : "";
    });
  });
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function getPersonRotaProfile_(employeeId) {
  const cleanId = String(employeeId || "").trim();
  const person = getStaffLiteByEmployeeId_(cleanId) || {};
  return {
    ok: true,
    employeeId: cleanId,
    person: person,
    standardRota: getPersonStandardRota_(cleanId)
  };
}

function markEmployeeSick_(employeeId, dateText, notes) {
  const cleanId = String(employeeId || "").trim();
  const date = normaliseWorkforceDate_(dateText) || Utilities.formatDate(new Date(), WORKFORCE_CONFIG.timeZone, "yyyy-MM-dd");
  const person = getStaffLiteByEmployeeId_(cleanId);
  if (!person || !person.name) throw new Error("Could not find that employee.");
  const spreadsheet = getWorkforceSpreadsheet_();
  const absenceId = ["manual_sick", cleanId || slugifyWorkforce_(person.name), date.replace(/-/g, "")].join("_");
  upsertWorkforceRows_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.absences), "Absence ID", [{
    "Absence ID": absenceId,
    "Employee ID": cleanId,
    "Employee Name": person.name,
    "Absence Type": "Sickness",
    "Start Date": date,
    "End Date": date,
    "Status": "Manual",
    "Source": "Rota Cockpit",
    "BrightHR Raw JSON": "",
    "Last Synced": new Date()
  }]);
  return {
    ok: true,
    employeeId: cleanId,
    employeeName: person.name,
    date: date,
    message: person.name + " marked as sick for " + date + ". Reloading suggestions."
  };
}

function getPeopleStandardRotaForWeek_(siteId, weekStartDate) {
  const cleanSiteId = String(siteId || "").trim();
  const weekStart = normaliseWeekStart_(weekStartDate);
  const weekDates = getWeekDates_(weekStart);
  const dateByWeekday = {};
  weekDates.forEach(function(day) {
    dateByWeekday[day.weekday] = day.date;
  });
  const spreadsheet = getWorkforceSpreadsheet_();
  return readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.personStandardRota))
    .filter(function(row) {
      if (!workforceBoolean_(row.Active)) return false;
      const weekday = String(row["Day of Week"] || "");
      const date = dateByWeekday[weekday];
      if (!date) return false;
      const from = normaliseWorkforceDate_(row["Effective From"]);
      const to = normaliseWorkforceDate_(row["Effective To"]);
      if (from && date < from) return false;
      if (to && date > to) return false;
      return !cleanSiteId || String(row.Site || "") === cleanSiteId || String(row.Site || "");
    });
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

function buildCockpitStaffLiteLookup_(spreadsheet) {
  const lookup = { byId: {}, byName: {} };
  const sheet = spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory);
  if (!sheet || sheet.getLastRow() < 2) return lookup;
  const map = workforceHeaderMap_(sheet);
  const rowCount = sheet.getLastRow() - 1;
  const idValues = map["Employee ID"] ? sheet.getRange(2, map["Employee ID"], rowCount, 1).getDisplayValues() : [];
  const nameValues = map.Name ? sheet.getRange(2, map.Name, rowCount, 1).getDisplayValues() : [];
  const emailValues = map.Email ? sheet.getRange(2, map.Email, rowCount, 1).getDisplayValues() : [];
  const roleValues = map.Role ? sheet.getRange(2, map.Role, rowCount, 1).getDisplayValues() : [];
  for (let index = 0; index < rowCount; index += 1) {
    const person = {
      employeeId: (idValues[index] && idValues[index][0]) || "",
      name: (nameValues[index] && nameValues[index][0]) || "",
      email: (emailValues[index] && emailValues[index][0]) || "",
      role: (roleValues[index] && roleValues[index][0]) || ""
    };
    if (person.employeeId) lookup.byId[person.employeeId] = person;
    if (person.name) lookup.byName[normaliseWorkforcePerson_(person.name)] = person;
  }
  return lookup;
}

function getSuggestionConfidence_(score) {
  if (score >= 90) return "High";
  if (score >= 55) return "Medium";
  return "Low";
}

function getRotaCockpitTotals_(days) {
  const totals = getEmptyRotaCockpitTotals_();
  (days || []).forEach(function(day) {
    totals.scheduled += (day.fullRota || day.scheduled || []).length;
    totals.absences += (day.absences || []).length;
    (day.issues || []).forEach(function(issue) {
      if (String(issue.type || "") === "conflict") totals.conflicts += 1;
      if (/resolved|snoozed|ignored/i.test(String(issue.status || ""))) {
        totals.resolved += 1;
      } else if (issue.bestSuggestion) {
        totals.coverSuggested += 1;
        totals.suggestions += 1;
      } else if (String(issue.type || "") === "uncovered") {
        totals.uncoveredGaps += 1;
        totals.uncovered += 1;
      } else {
        totals.unresolvedActions += 1;
      }
    });
  });
  return totals;
}

function getEmptyRotaCockpitTotals_() {
  return {
    scheduled: 0,
    absences: 0,
    uncovered: 0,
    uncoveredGaps: 0,
    suggestions: 0,
    coverSuggested: 0,
    conflicts: 0,
    unresolvedActions: 0,
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

function getBrightHrAdminSettings_() {
  const keys = WORKFORCE_CONFIG.scriptProperties;
  const properties = PropertiesService.getScriptProperties();
  return {
    ok: true,
    apiBaseUrl: properties.getProperty(keys.brightHrApiBaseUrl) || WORKFORCE_CONFIG.brightHr.apiBaseUrl,
    employeeEndpoint: properties.getProperty(keys.brightHrEmployeesPath) || "employees/v1/query",
    employeeMethod: properties.getProperty(keys.brightHrEmployeesMethod) || "post",
    absenceEndpoint: properties.getProperty(keys.brightHrAbsencesPath) || "",
    absenceMethod: properties.getProperty(keys.brightHrAbsencesMethod) || "post",
    tokenPropertyKey: properties.getProperty(keys.brightHrTokenPropertyKey) || keys.brightHrAccessToken,
    absenceLookaheadDays: properties.getProperty(keys.brightHrAbsenceLookaheadDays) || "28",
    absenceLookbackDays: properties.getProperty(keys.brightHrAbsenceLookbackDays) || "3"
  };
}

function saveBrightHrAdminSettings_(settings) {
  const keys = WORKFORCE_CONFIG.scriptProperties;
  const properties = PropertiesService.getScriptProperties();
  const apiBaseUrl = String(settings.apiBaseUrl || "").trim().replace(/\/+$/, "");
  const employeeEndpoint = String(settings.employeeEndpoint || "employees/v1/query").trim().replace(/^\/+/, "");
  const employeeMethod = String(settings.employeeMethod || "post").trim().toLowerCase();
  const absenceEndpoint = String(settings.absenceEndpoint || "").trim().replace(/^\/+/, "");
  const absenceMethod = String(settings.absenceMethod || "post").trim().toLowerCase();
  if (apiBaseUrl) properties.setProperty(keys.brightHrApiBaseUrl, apiBaseUrl);
  if (employeeEndpoint) properties.setProperty(keys.brightHrEmployeesPath, employeeEndpoint);
  if (["get", "post"].indexOf(employeeMethod) !== -1) properties.setProperty(keys.brightHrEmployeesMethod, employeeMethod);
  if (absenceEndpoint) properties.setProperty(keys.brightHrAbsencesPath, absenceEndpoint);
  if (["get", "post"].indexOf(absenceMethod) !== -1) properties.setProperty(keys.brightHrAbsencesMethod, absenceMethod);
  properties.setProperty(keys.brightHrTokenPropertyKey, String(settings.tokenPropertyKey || keys.brightHrAccessToken).trim());
  const cappedRange = capBrightHrAbsenceRange_(settings.absenceLookbackDays, settings.absenceLookaheadDays);
  properties.setProperty(keys.brightHrAbsenceLookaheadDays, String(cappedRange.lookahead));
  properties.setProperty(keys.brightHrAbsenceLookbackDays, String(cappedRange.lookback));
  return {
    ok: true,
    message: "BrightHR admin settings saved."
  };
}

function testBrightHrEndpointPreview_(type) {
  const settings = getBrightHrAdminSettings_();
  const isEmployee = type === "employee";
  const endpoint = isEmployee ? settings.employeeEndpoint : settings.absenceEndpoint;
  const method = isEmployee ? settings.employeeMethod : settings.absenceMethod;
  if (!endpoint) {
    return {
      ok: false,
      endpoint: "",
      method: method,
      message: "No " + (isEmployee ? "employee" : "absence") + " endpoint is configured."
    };
  }
  try {
    const data = brightHrRequest_(endpoint, method === "post" ? {
      method: "post",
      payload: buildBrightHrPreviewPayload_(isEmployee)
    } : { method: "get" });
    return {
      ok: true,
      endpoint: endpoint,
      method: method,
      itemCount: getBrightHrItems_(data).length,
      rawPreview: JSON.stringify(data).slice(0, 1800),
      message: "BrightHR " + (isEmployee ? "employee" : "absence") + " endpoint test worked."
    };
  } catch (error) {
    return {
      ok: false,
      endpoint: endpoint,
      method: method,
      httpStatus: error.httpStatus || "",
      rawPreview: String(error.responseText || "").slice(0, 1800),
      error: error.message || String(error),
      message: "BrightHR " + (isEmployee ? "employee" : "absence") + " endpoint test failed: " + (error.message || String(error))
    };
  }
}

function buildBrightHrPreviewPayload_(isEmployee) {
  const settings = getBrightHrAdminSettings_();
  if (isEmployee) return { pageSize: 5 };
  const employeeId = getFirstActiveEmployeeIdForBrightHrTest_();
  if (!employeeId) {
    throw new Error("No employee ID found in Staff Directory. Sync employees before testing the absence endpoint.");
  }
  const range = capBrightHrAbsenceRange_(settings.absenceLookbackDays, settings.absenceLookaheadDays);
  const today = new Date();
  const from = new Date(today.getFullYear(), today.getMonth(), today.getDate() - range.lookback);
  const to = new Date(today.getFullYear(), today.getMonth(), today.getDate() + range.lookahead);
  return {
    filters: {
      employeeId: employeeId,
      from: formatDateInputForBrightHr_(from),
      to: formatDateInputForBrightHr_(to)
    },
    pageSize: 5
  };
}

function formatDateInputForBrightHr_(date) {
  return Utilities.formatDate(date, WORKFORCE_CONFIG.timeZone, "yyyy-MM-dd");
}

function capBrightHrAbsenceRange_(lookbackValue, lookaheadValue) {
  const requestedLookback = Math.max(0, Number(lookbackValue || 3));
  const requestedLookahead = Math.max(1, Number(lookaheadValue || 28));
  const lookback = Math.min(requestedLookback, 30);
  const lookahead = Math.min(requestedLookahead, Math.max(1, 31 - lookback));
  return {
    lookback: lookback,
    lookahead: lookahead
  };
}

function getFirstActiveEmployeeIdForBrightHrTest_() {
  const spreadsheet = getWorkforceSpreadsheet_();
  const person = readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory))
    .filter(function(row) {
      return String(row["Employee ID"] || "").trim() &&
        String(row["Employment Status"] || "").toLowerCase() !== "terminated" &&
        !workforceBoolean_(row.Terminated);
    })[0];
  return person ? String(person["Employee ID"] || "").trim() : "";
}
