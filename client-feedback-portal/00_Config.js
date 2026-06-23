const FEEDBACK_CONFIG = Object.freeze({
  timeZone: "Europe/London",
  requestDelayHours: 24,
  requestExpiryHours: 72,
  eligibleStatuses: ["CONFIRMED"],
  sheets: {
    requests: "Feedback Requests",
    responses: "Feedback Responses",
    itemRatings: "Feedback Item Ratings",
    settings: "Feedback Settings"
  },
  branding: {
    fikaLogoUrl: "",
    fikaFallbackText: "Fika",
    accent: "#3d21bf",
    ink: "#221874",
    paper: "#f4f4f2"
  }
});

/**
 * Add future sites here. Spreadsheet IDs deliberately live in Script
 * Properties, so cloning a dashboard never requires committing credentials.
 */
const FEEDBACK_SITES = Object.freeze({
  angel_court: Object.freeze({
    siteId: "angel_court",
    siteName: "Angel Court",
    clientFacingName: "Angel Court Hospitality Feedback",
    dashboardSheetName: "Dashboard Data",
    spreadsheetPropertyKey: "ANGEL_COURT_DASHBOARD_SPREADSHEET_ID",
    eligibleStatuses: ["CONFIRMED"],
    requestDelayHours: 24,
    requestExpiryHours: 72,
    siteFallbackText: "Angel Court"
  })
});

const FEEDBACK_SETTINGS = Object.freeze([
  { key: "FEEDBACK_WEB_APP_URL", fallback: "" },
  { key: "SITE_EMAIL_ADDRESS", fallback: "seven@fikacatering.com" },
  { key: "FOLLOW_UP_RECIPIENTS", fallback: "" },
  { key: "FIKA_LOGO_URL", fallback: FEEDBACK_CONFIG.branding.fikaLogoUrl },
  { key: "SITE_LOGO_URL", fallback: "" },
  { key: "FIKA_FALLBACK_TEXT", fallback: FEEDBACK_CONFIG.branding.fikaFallbackText },
  { key: "SITE_FALLBACK_TEXT", fallback: "" },
  { key: "COLOUR_ACCENT", fallback: FEEDBACK_CONFIG.branding.accent },
  { key: "COLOUR_INK", fallback: FEEDBACK_CONFIG.branding.ink },
  { key: "COLOUR_PAPER", fallback: FEEDBACK_CONFIG.branding.paper }
]);

function getFeedbackSite_(siteId) {
  const site = FEEDBACK_SITES[String(siteId || "").trim()];
  if (!site) throw new Error("Unknown feedback site: " + siteId);
  return site;
}

function getFeedbackRegistryStatus() {
  const properties = PropertiesService.getScriptProperties();
  return Object.keys(FEEDBACK_SITES).map(function(siteId) {
    const site = FEEDBACK_SITES[siteId];
    const configured = Boolean(
      properties.getProperty(site.spreadsheetPropertyKey) ||
      (siteId === "angel_court" && properties.getProperty("DASHBOARD_SPREADSHEET_ID"))
    );
    return {
      siteId: site.siteId,
      siteName: site.siteName,
      propertyKey: site.spreadsheetPropertyKey,
      configured: configured
    };
  });
}

function getFeedbackSiteContexts_() {
  const properties = PropertiesService.getScriptProperties();
  return Object.keys(FEEDBACK_SITES).map(function(siteId) {
    const site = FEEDBACK_SITES[siteId];
    let spreadsheetId = properties.getProperty(site.spreadsheetPropertyKey);
    if (!spreadsheetId && siteId === "angel_court") {
      spreadsheetId = properties.getProperty("DASHBOARD_SPREADSHEET_ID");
    }
    if (!spreadsheetId) return null;
    return {
      site: site,
      spreadsheetId: spreadsheetId,
      spreadsheet: SpreadsheetApp.openById(spreadsheetId)
    };
  }).filter(Boolean);
}

function getFeedbackSiteContext_(siteId) {
  const site = getFeedbackSite_(siteId);
  const contexts = getFeedbackSiteContexts_();
  for (let index = 0; index < contexts.length; index++) {
    if (contexts[index].site.siteId === site.siteId) return contexts[index];
  }
  throw new Error(
    "No spreadsheet is configured for " + site.siteName +
    ". Set Script Property " + site.spreadsheetPropertyKey + "."
  );
}

function getFeedbackSettings_(context) {
  const site = context.site;
  const values = {};
  FEEDBACK_SETTINGS.forEach(function(setting) {
    values[setting.key] = setting.fallback;
  });
  values.SITE_FALLBACK_TEXT = site.siteFallbackText || site.siteName;
  try {
    const sheet = context.spreadsheet.getSheetByName(FEEDBACK_CONFIG.sheets.settings);
    if (!sheet || sheet.getLastRow() < 2) return values;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2)
      .getDisplayValues()
      .forEach(function(row) {
        const key = String(row[0] || "").trim();
        const value = String(row[1] || "").trim();
        if (value && Object.prototype.hasOwnProperty.call(values, key)) {
          values[key] = value;
        }
      });
  } catch (error) {
    console.warn("Feedback settings could not be read: " + error.message);
  }
  return values;
}

function getPublicFeedbackConfig_(token) {
  const request = token ? findFeedbackRequestByToken_(token) : null;
  const context = request ? request.context : getFeedbackSiteContexts_()[0];
  const site = context ? context.site : {
    siteName: "Hospitality",
    clientFacingName: "FIKA Hospitality Feedback",
    siteFallbackText: "Hospitality"
  };
  const settings = context ? getFeedbackSettings_(context) : {};
  return {
    siteId: site.siteId || "",
    siteName: site.siteName,
    clientFacingName: site.clientFacingName,
    branding: {
      fikaLogoUrl: settings.FIKA_LOGO_URL || "",
      fikaFallbackText: settings.FIKA_FALLBACK_TEXT || "Fika",
      siteLogoUrl: settings.SITE_LOGO_URL || "",
      siteFallbackText: settings.SITE_FALLBACK_TEXT || site.siteFallbackText,
      accent: settings.COLOUR_ACCENT || FEEDBACK_CONFIG.branding.accent,
      ink: settings.COLOUR_INK || FEEDBACK_CONFIG.branding.ink,
      paper: settings.COLOUR_PAPER || FEEDBACK_CONFIG.branding.paper
    }
  };
}
