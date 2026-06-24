const WORKFORCE_CONFIG = Object.freeze({
  appName: "FIKA Workforce Operations Platform",
  appVersion: "v0.5.2",
  timeZone: "Europe/London",
  hrProvider: "brighthr",
  brightHr: {
    tokenUrl: "https://login.brighthr.com/connect/token",
    apiBaseUrl: "https://api.bright.hr"
  },
  scriptProperties: {
    workforceSpreadsheetId: "WORKFORCE_SPREADSHEET_ID",
    legacyRotaSpreadsheetId: "LEGACY_ROTA_SPREADSHEET_ID",
    reliefRotaSpreadsheetId: "RELIEF_ROTA_SPREADSHEET_ID",
    brightHrClientId: "BRIGHTHR_CLIENT_ID",
    brightHrClientSecret: "BRIGHTHR_CLIENT_SECRET",
    brightHrTokenUrl: "BRIGHTHR_TOKEN_URL",
    brightHrApiBaseUrl: "BRIGHTHR_API_BASE_URL",
    brightHrAbsencesPath: "BRIGHTHR_ABSENCES_PATH",
    brightHrAccessToken: "BRIGHTHR_ACCESS_TOKEN",
    brightHrAccessTokenExpiresAt: "BRIGHTHR_ACCESS_TOKEN_EXPIRES_AT"
  },
  sheets: {
    staffDirectory: "Staff Directory",
    absences: "Absences",
    sites: "Sites",
    managers: "Managers",
    agencyContacts: "Agency Contacts",
    emailTemplates: "Email Templates",
    roleLibrary: "Role Library",
    shiftPatterns: "Shift Patterns",
    notificationRules: "Notification Rules",
    costRates: "Cost Rates",
    siteRoles: "Site Roles",
    rotaTemplates: "Rota Templates",
    rotaExceptions: "Rota Exceptions",
    rotaShifts: "Rota Shifts",
    reliefSuggestions: "Relief Suggestions",
    agencyRequests: "Agency Requests",
    coverageGaps: "Coverage Gaps",
    reliefAvailability: "Relief Availability",
    settings: "Workforce Settings"
  },
  legacyRota: {
    sourceName: "Legacy Rota Import",
    importStartDate: "2026-06-29",
    dayAssignmentColumns: [2, 4, 6, 8, 10, 12, 14],
    firstDataRow: 4
  },
  reliefRota: {
    sourceName: "Relief Rota Import",
    importStartDate: "2026-06-29",
    dayColumns: [2, 3, 4, 5, 6, 7, 8]
  }
});
