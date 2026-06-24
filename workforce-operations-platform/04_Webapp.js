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
      options: options
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
      options: { sites: [], defaultWeekStart: "" }
    }));
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
    result.warnings.push(error.message || String(error));
  }
  result.message = [
    "Sync complete.",
    result.employees ? (result.employees.synced || 0) + " employee(s) updated." : "",
    result.absences ? (result.absences.synced || 0) + " absence row(s) updated." : "",
    result.warnings.length ? "Absences need attention: " + result.warnings.join(" ") : ""
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
  return getWeeklyRotaBoard(siteId, weekStartDate);
}

function generateReliefSuggestionsFromUi() {
  return generateReliefSuggestions(28);
}

function createAgencyRequestForGapFromUi(gapId, agencyId) {
  return createAgencyRequestForGap(gapId, agencyId);
}

function markCoverageGapFromUi(gapId, action, details) {
  return markCoverageGap_(gapId, action, details || {});
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

function safeGetCoverageGapSummary_() {
  try {
    return getCoverageGapSummary(8);
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
      modules: []
    };
  }
}

function getWorkforceSummary_(spreadsheet) {
  return {
    staffCount: getActiveStaffCount_(spreadsheet),
    absenceCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.absences),
    siteCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.sites),
    rotaShiftCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.rotaShifts),
    reliefSuggestionCount: getWorkforceSheetRecordCount_(spreadsheet, WORKFORCE_CONFIG.sheets.reliefSuggestions),
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
  return readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory))
    .filter(function(person) {
      const employmentStatus = String(person["Employment Status"] || "").toLowerCase();
      return String(person.Name || "").trim() &&
        employmentStatus !== "terminated" &&
        !workforceBoolean_(person.Terminated);
    }).length;
}
