const SITE_CONFIG = Object.freeze({
  siteId: "angel_court",
  siteName: "Angel Court",
  address: "1 Angel Court, London",
  clientFacingName: "Angel Court Hospitality",
  currency: "GBP",
  locale: "en-GB",
  timeZone: "Europe/London",
  bookingReferencePrefix: "AC",
  siteEmailAddress: "seven@fikacatering.com",
  notificationRecipients: [],
  integration: {
    mode: "direct_dashboard",
    spreadsheetId: "",
    dashboardSheetName: "Dashboard Data",
    managementFeeRate: 0.08,
    vatRate: 0.20,
    keepClientRequestLog: false
  },
  sheets: {
    bookingRequests: "Booking Requests",
    bookingLineItems: "Booking Line Items",
    menuItems: "Menu Items",
    settings: "Platform Settings"
  },
  rules: {
    standardNoticeHours: 72,
    largeEventNoticeWorkingDays: 10,
    dietaryNoticeWorkingDays: 3,
    blockMinimumOrderIssues: true
  },
  copy: {
    vatNote: "Estimated total, subject to final confirmation, VAT, labour and equipment hire where applicable.",
    requestAcknowledgement: "I understand this is a booking request and is subject to confirmation."
  },
  branding: {
    eyebrow: "FIKA hospitality",
    heroTitle: "A fresh way to book hospitality.",
    heroBody: "Plan food, drinks and event service for Angel Court in one clear request.",
    fikaLogoUrl: "",
    fikaLogoAlt: "FIKA",
    fikaFallbackText: "Fika",
    siteLogoUrl: "",
    siteLogoAlt: "Angel Court",
    siteFallbackText: "Angel Court",
    accent: "#3d21bf",
    ink: "#221874",
    paper: "#f4f4f2"
  }
});

const PLATFORM_SETTINGS_SCHEMA = Object.freeze([
  { key: "FIKA_LOGO_URL", fallback: SITE_CONFIG.branding.fikaLogoUrl },
  { key: "FIKA_LOGO_ALT", fallback: SITE_CONFIG.branding.fikaLogoAlt },
  { key: "FIKA_FALLBACK_TEXT", fallback: SITE_CONFIG.branding.fikaFallbackText },
  { key: "SITE_LOGO_URL", fallback: SITE_CONFIG.branding.siteLogoUrl },
  { key: "SITE_LOGO_ALT", fallback: SITE_CONFIG.branding.siteLogoAlt },
  { key: "SITE_FALLBACK_TEXT", fallback: SITE_CONFIG.branding.siteFallbackText },
  { key: "CLIENT_FACING_NAME", fallback: SITE_CONFIG.clientFacingName },
  { key: "BRAND_EYEBROW", fallback: SITE_CONFIG.branding.eyebrow },
  { key: "HERO_TITLE", fallback: SITE_CONFIG.branding.heroTitle },
  { key: "HERO_BODY", fallback: SITE_CONFIG.branding.heroBody },
  { key: "COLOUR_ACCENT", fallback: SITE_CONFIG.branding.accent },
  { key: "COLOUR_INK", fallback: SITE_CONFIG.branding.ink },
  { key: "COLOUR_PAPER", fallback: SITE_CONFIG.branding.paper },
  { key: "SITE_EMAIL_ADDRESS", fallback: SITE_CONFIG.siteEmailAddress },
  { key: "NOTIFICATION_RECIPIENTS", fallback: SITE_CONFIG.notificationRecipients.join(", ") },
  { key: "DASHBOARD_URL", fallback: "" }
]);

const EVENT_TYPES = Object.freeze([
  { id: "breakfast", label: "Breakfast", categories: ["Breakfast"], noticeType: "standard" },
  { id: "lunch", label: "Lunch", categories: ["Lunch", "Lunch Add-ons", "Add-ons"], noticeType: "standard" },
  { id: "meeting_hospitality", label: "Meeting hospitality", categories: ["Breakfast", "Lunch", "Lunch Add-ons", "Drinks & Packages"], noticeType: "standard" },
  { id: "event_catering", label: "Event catering", categories: ["Events Catering", "Add-ons", "White Wine", "Red Wine", "Rosé Wine", "Sparkling Wine", "Drinks & Packages"], noticeType: "large" },
  { id: "drinks_reception", label: "Drinks reception", categories: ["White Wine", "Red Wine", "Rosé Wine", "Sparkling Wine", "Drinks & Packages", "Add-ons"], noticeType: "large" },
  { id: "bbq", label: "BBQ / large event", categories: ["Summer BBQ Catering", "White Wine", "Red Wine", "Rosé Wine", "Sparkling Wine", "Drinks & Packages"], noticeType: "large" },
  { id: "bespoke", label: "Bespoke enquiry", categories: [], noticeType: "large", allowsEmptyOrder: true }
]);

function getPublicPlatformConfig() {
  const settings = getPlatformSettings_();
  return {
    site: buildPublicSiteConfig_(settings),
    eventTypes: EVENT_TYPES,
    menu: MENU_SCHEMA.filter(function(item) { return item.available; })
  };
}

function buildPublicSiteConfig_(settings) {
  return Object.assign({}, SITE_CONFIG, {
    clientFacingName: settings.CLIENT_FACING_NAME,
    branding: Object.assign({}, SITE_CONFIG.branding, {
      eyebrow: settings.BRAND_EYEBROW,
      heroTitle: settings.HERO_TITLE,
      heroBody: settings.HERO_BODY,
      fikaLogoUrl: settings.FIKA_LOGO_URL,
      fikaLogoAlt: settings.FIKA_LOGO_ALT,
      fikaFallbackText: settings.FIKA_FALLBACK_TEXT,
      siteLogoUrl: settings.SITE_LOGO_URL,
      siteLogoAlt: settings.SITE_LOGO_ALT,
      siteFallbackText: settings.SITE_FALLBACK_TEXT,
      accent: settings.COLOUR_ACCENT,
      ink: settings.COLOUR_INK,
      paper: settings.COLOUR_PAPER
    })
  });
}

function getPlatformSettings_() {
  const values = {};
  PLATFORM_SETTINGS_SCHEMA.forEach(function(setting) { values[setting.key] = setting.fallback; });

  try {
    const spreadsheet = getBookingSpreadsheet_();
    const sheet = spreadsheet.getSheetByName(SITE_CONFIG.sheets.settings);
    if (!sheet || sheet.getLastRow() < 2) return values;

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
    rows.forEach(function(row) {
      const key = String(row[0] || "").trim();
      if (Object.prototype.hasOwnProperty.call(values, key) && String(row[1] || "").trim()) {
        values[key] = String(row[1]).trim();
      }
    });
  } catch (error) {
    console.warn("Platform Settings could not be read: " + error.message);
  }

  return values;
}
