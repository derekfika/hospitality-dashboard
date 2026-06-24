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
  result.message = [
    "Employee sync complete.",
    result.employees ? (result.employees.synced || 0) + " employee(s) updated." : "",
    "Run absence sync, gap detection and relief planning separately."
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

function requestAgencyForGapFromUi(gapId, agencyId, details) {
  return requestAgencyForGap_(gapId, agencyId, details || {});
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
