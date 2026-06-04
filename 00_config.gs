const CONFIG = {
  QUOTE_TEMPLATE_DOC_ID: "1-2E0p2D7ivEyL7iB7yTe3tsgoppwSq48itHw7fSaBl0",
  QUOTE_ROOT_FOLDER_NAME: "Hospitality",
  SHEET_NAME: "Dashboard Data",

  PRINTER_EMAIL: "mvt5149ww1d8q0@print.epsonconnect.com",

  DRIVE: {
    ROOT_FOLDER_NAME: "Hospitality"
  },

  CALENDAR: {
    TEST_CALENDAR: "derek@fikacatering.com"
  },

  STATUS: {
    NEW: "NEW",
    NEEDS_REVIEW: "NEEDS_REVIEW",
    READY: "READY",
    QUOTE_GENERATED: "QUOTE_GENERATED",
    PRINTED: "PRINTED",
    CPU_CREATED: "CPU_CREATED",
    CONFIRMED: "CONFIRMED",
    CANCELLED: "CANCELLED",
    ERROR: "ERROR",
    ARCHIVED: "ARCHIVED"
  },

  CALENDAR_ID: "seven@fikacatering.com",

  CALENDAR_ATTENDEES: ["seven@fikacatering.com",
    "derek@fikacatering.com",
    // "cpu@fikacatering.com"
  ], CALENDAR_EVENT_COLOR_ID: "9", // blue / blueberry
  CALENDAR_EVENT_DURATION_MINUTES: 60,

  // ---------- APP / LOCATION ----------
APP_NAME: "Angel Court Hospitality Dashboard",
LOCATION_NAME: "Angel Court",
LOCATION_SHORT_CODE: "OAC",

// ---------- BRANDING ----------
FIKA_LOGO_URL: "https://bloom-coffee.org/bloom-quiz/assets/images/fika_logoRGB.png",
ANGEL_COURT_LOGO_URL: "",
SEVEN_LOGO_URL: "",
FAVICON_URL: "",

// ---------- FONTS ----------
FONT_TITLE_URL: "",
FONT_BODY_URL: "",
FONT_TITLE_FAMILY: "VIM",
FONT_BODY_FAMILY: "Gilroy",

// ---------- UI COLOURS ----------
COLOUR_BACKGROUND: "#F8F6F2",
COLOUR_SURFACE: "#FFFFFF",
COLOUR_SURFACE_ALT: "#F1EEE8",

COLOUR_TEXT: "#252525",
COLOUR_TEXT_MUTED: "#777777",
COLOUR_BORDER: "#DED8CD",

COLOUR_PRIMARY: "#202020",
COLOUR_SECONDARY: "#8A8A8A",

COLOUR_SUCCESS: "#6C8A74",
COLOUR_WARNING: "#B08B52",
COLOUR_DANGER: "#8B5E5E",
COLOUR_INFO: "#5B7EA0",

};

function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Settings");

  if (!sh) {
    sh = ss.insertSheet("Settings");
    sh.getRange(1, 1, 1, 2).setValues([["Key", "Value"]]);
  }

  return sh;
}

function getSettings_() {
  const sh = getSettingsSheet_();
  const values = sh.getDataRange().getValues();
  const settings = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || "").trim();
    let value = values[i][1];

    if (!key) continue;

    if (typeof value === "string") {
      const upper = value.trim().toUpperCase();

      if (upper === "TRUE") value = true;
      else if (upper === "FALSE") value = false;
    }

    settings[key] = value;
  }

  return settings;
}

function getSetting_(key, fallback) {
  const sh = getSettingsSheet_();
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      return values[i][1] || fallback;
    }
  }

  return fallback;
}

function setSetting_(key, value) {
  const sh = getSettingsSheet_();
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  sh.appendRow([key, value]);
}


const SETTINGS = getSettings_();

Logger.log(SETTINGS.SITE_NAME);
Logger.log(SETTINGS.MGMT_FEE_PERCENT);
Logger.log(SETTINGS.CALENDAR_ID);