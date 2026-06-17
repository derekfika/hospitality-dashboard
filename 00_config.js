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
    CPU_CREATED: "CPU_CREATED",
    CONFIRMED: "CONFIRMED",
    CANCELLED: "CANCELLED",
    ERROR: "ERROR",
    ARCHIVED: "ARCHIVED"
  },

  CALENDAR_ID: "seven@fikacatering.com",

  CALENDAR_ATTENDEES: ["cpux@fikacatering.com, dwayne@fikacatering.com, logistics@fikacatering.com"
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

COLOUR_PRIMARY: "#4F34C7",
COLOUR_SECONDARY: "#280F8C",

COLOUR_SUCCESS: "#6C8A74",
COLOUR_WARNING: "#B08B52",
COLOUR_DANGER: "#8B5E5E",
COLOUR_INFO: "#5B7EA0",

};

function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Settings");
  const headers = ["Key", "Value", "Section", "Label", "Type", "Notes"];

  if (!sh) {
    sh = ss.insertSheet("Settings");
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (sh.getLastColumn() < headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
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

function getSettingSchema_() {
  return [
    {
      id: "general",
      title: "General",
      fields: [
        { key: "APP_NAME", label: "App name", type: "text", fallback: CONFIG.APP_NAME, required: true, notes: "Shown in the browser title and dashboard header." },
        { key: "LOCATION_NAME", label: "Location name", type: "text", fallback: CONFIG.LOCATION_NAME, notes: "Human-friendly site name." },
        { key: "LOCATION_SHORT_CODE", label: "Location short code", type: "text", fallback: CONFIG.LOCATION_SHORT_CODE, required: true, notes: "Used at the start of calendar event titles." },
        { key: "SHEET_NAME", label: "Dashboard sheet name", type: "text", fallback: CONFIG.SHEET_NAME, required: true, notes: "Must match the tab that stores dashboard booking rows." }
      ]
    },
    {
      id: "inbox",
      title: "Inbox Scan",
      fields: [
        { key: "PROCESSED_LABEL_NAME", label: "Processed Gmail label", type: "text", fallback: "AC_HOSPITALITY_PROCESSED", required: true, notes: "Threads with this Gmail label are excluded from future scans." },
        { key: "EARLIEST_SCAN_DATE", label: "Earliest scan date", type: "date", fallback: "", notes: "Optional. Use YYYY-MM-DD to stop scans going further back." },
        { key: "LAST_INBOX_SCAN_AT", label: "Last scan timestamp", type: "text", fallback: "", readonly: true, notes: "Updated automatically after a completed inbox scan." },
        { key: "ARCHIVE_AFTER_DAYS", label: "Archive after days", type: "number", fallback: 0, min: 0, notes: "Bookings older than this can be archived automatically. Use 0 to archive before today." }
      ]
    },
    {
      id: "quotes",
      title: "Quotes",
      fields: [
        { key: "QUOTE_TEMPLATE_DOC_ID", label: "Quote template document ID", type: "text", fallback: CONFIG.QUOTE_TEMPLATE_DOC_ID, required: true, notes: "Google Doc template ID used when generating quote files." },
        { key: "QUOTE_ROOT_FOLDER_NAME", label: "Quote root folder", type: "text", fallback: CONFIG.QUOTE_ROOT_FOLDER_NAME, required: true, notes: "Drive folder name where quote folders are created." },
        { key: "PRINTER_EMAIL", label: "Printer email", type: "email", fallback: CONFIG.PRINTER_EMAIL, notes: "Email address used by the print quote action." }
      ]
    },
    {
      id: "calendar",
      title: "Calendar",
      fields: [
        { key: "CALENDAR_ID", label: "Calendar ID", type: "text", fallback: CONFIG.CALENDAR_ID, required: true, notes: "Calendar email or ID where CPU events are created." },
        { key: "CALENDAR_ATTENDEES", label: "Calendar attendees", type: "textarea", fallback: (CONFIG.CALENDAR_ATTENDEES || []).join(", "), notes: "Comma or line separated email addresses invited to CPU events." },
        { key: "CALENDAR_EVENT_COLOR_ID", label: "Event colour ID", type: "text", fallback: CONFIG.CALENDAR_EVENT_COLOR_ID, notes: "Google Calendar colour ID. Usually a number from 1 to 11." },
        { key: "CALENDAR_EVENT_DURATION_MINUTES", label: "Event duration minutes", type: "number", fallback: CONFIG.CALENDAR_EVENT_DURATION_MINUTES, required: true, min: 1, notes: "Length of created calendar events." }
      ]
    },
    {
      id: "branding",
      title: "Branding",
      fields: [
        { key: "FIKA_LOGO_URL", label: "FIKA logo URL", type: "url", fallback: CONFIG.FIKA_LOGO_URL, notes: "Main logo shown in the dashboard header." },
        { key: "ANGEL_COURT_LOGO_URL", label: "Angel Court logo URL", type: "url", fallback: CONFIG.ANGEL_COURT_LOGO_URL, notes: "Reserved for future branded quote/dashboard use." },
        { key: "SEVEN_LOGO_URL", label: "Seven logo URL", type: "url", fallback: CONFIG.SEVEN_LOGO_URL, notes: "Reserved for future branded quote/dashboard use." },
        { key: "FAVICON_URL", label: "Favicon URL", type: "url", fallback: CONFIG.FAVICON_URL, notes: "Optional browser tab icon." },
        { key: "FONT_TITLE_URL", label: "Title font URL", type: "url", fallback: CONFIG.FONT_TITLE_URL, notes: "Optional font file URL for headings." },
        { key: "FONT_BODY_URL", label: "Body font URL", type: "url", fallback: CONFIG.FONT_BODY_URL, notes: "Optional font file URL for body text." },
        { key: "FONT_TITLE_FAMILY", label: "Title font family", type: "text", fallback: CONFIG.FONT_TITLE_FAMILY, required: true, notes: "CSS font-family name for headings." },
        { key: "FONT_BODY_FAMILY", label: "Body font family", type: "text", fallback: CONFIG.FONT_BODY_FAMILY, required: true, notes: "CSS font-family name for controls and table text." }
      ]
    },
    {
      id: "colours",
      title: "Colours",
      fields: [
        { key: "COLOUR_BACKGROUND", label: "Background", type: "color", fallback: CONFIG.COLOUR_BACKGROUND, required: true, notes: "Hex colour, for example #F8F6F2." },
        { key: "COLOUR_SURFACE", label: "Surface", type: "color", fallback: CONFIG.COLOUR_SURFACE, required: true, notes: "Hex colour used for raised panels." },
        { key: "COLOUR_SURFACE_ALT", label: "Alternate surface", type: "color", fallback: CONFIG.COLOUR_SURFACE_ALT, required: true, notes: "Hex colour used for card/table surfaces." },
        { key: "COLOUR_TEXT", label: "Text", type: "color", fallback: CONFIG.COLOUR_TEXT, required: true, notes: "Main text hex colour." },
        { key: "COLOUR_TEXT_MUTED", label: "Muted text", type: "color", fallback: CONFIG.COLOUR_TEXT_MUTED, required: true, notes: "Secondary text hex colour." },
        { key: "COLOUR_BORDER", label: "Border", type: "color", fallback: CONFIG.COLOUR_BORDER, required: true, notes: "Border hex colour." },
        { key: "COLOUR_PRIMARY", label: "Primary", type: "color", fallback: CONFIG.COLOUR_PRIMARY, required: true, notes: "Primary action and heading hex colour." },
        { key: "COLOUR_SECONDARY", label: "Secondary", type: "color", fallback: CONFIG.COLOUR_SECONDARY, required: true, notes: "Reserved secondary brand hex colour." },
        { key: "COLOUR_SUCCESS", label: "Success", type: "color", fallback: CONFIG.COLOUR_SUCCESS, required: true, notes: "Success feedback hex colour." },
        { key: "COLOUR_WARNING", label: "Warning", type: "color", fallback: CONFIG.COLOUR_WARNING, required: true, notes: "Warning feedback hex colour." },
        { key: "COLOUR_DANGER", label: "Danger", type: "color", fallback: CONFIG.COLOUR_DANGER, required: true, notes: "Danger/cancel feedback hex colour." },
        { key: "COLOUR_INFO", label: "Info", type: "color", fallback: CONFIG.COLOUR_INFO, required: true, notes: "Informational accent hex colour." }
      ]
    }
  ];
}

function getDashboardSettings() {
  ensureSettingsDefaults_();

  const stored = getSettings_();
  const schema = getSettingSchema_();
  const values = {};

  schema.forEach(section => {
    section.fields.forEach(field => {
      values[field.key] =
        stored[field.key] !== undefined && stored[field.key] !== ""
          ? stored[field.key]
          : field.fallback;
    });
  });

  return {
    ok: true,
    schema,
    sections: schema,
    values,
    validation: validateDashboardSettings_(values, schema)
  };
}

function ensureSettingsDefaults_() {
  const sh = getSettingsSheet_();
  const schema = getSettingSchema_();
  const values = sh.getDataRange().getValues();
  const existing = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || "").trim();
    if (key) existing[key] = i + 1;
  }

  const rowsToAppend = [];

  schema.forEach(section => {
    section.fields.forEach(field => {
      const metadata = [
        section.title,
        field.label,
        field.type || "text",
        field.notes || (field.readonly ? "Read only" : "")
      ];

      if (existing[field.key]) {
        sh.getRange(existing[field.key], 3, 1, metadata.length).setValues([metadata]);
        return;
      }

      rowsToAppend.push([
        field.key,
        normaliseSettingValue_(field.fallback, field),
        ...metadata
      ]);
    });
  });

  if (rowsToAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
  }
}

function saveDashboardSettings(values) {
  if (!values || typeof values !== "object") {
    throw new Error("No settings were supplied.");
  }

  const schema = getSettingSchema_();
  const editableFields = getEditableSettingFields_();
  const mergedValues = getDashboardSettings().values;
  const changedValues = {};

  Object.keys(values).forEach(key => {
    if (!editableFields[key]) return;
    changedValues[key] = normaliseSettingValue_(values[key], editableFields[key]);
    mergedValues[key] = changedValues[key];
  });

  const validation = validateDashboardSettings_(mergedValues, schema);
  if (!validation.ok) {
    throw new Error("Settings were not saved: " + validation.errors.join(" "));
  }

  Object.keys(changedValues).forEach(key => {
    setSetting_(key, changedValues[key]);
  });

  return getDashboardSettings();
}

function importDashboardSettingsDraft(sheetName) {
  sheetName = sheetName || "Settings Draft";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draftSheet = ss.getSheetByName(sheetName);

  if (!draftSheet) {
    throw new Error(`Settings draft sheet '${sheetName}' not found.`);
  }

  const values = draftSheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error(`Settings draft sheet '${sheetName}' has no values to import.`);
  }

  const headers = values[0].map(header => String(header || "").trim().toLowerCase());
  const keyIndex = findSettingsDraftColumn_(headers, [
    "key",
    "setting",
    "setting key",
    "setting_key",
    "name"
  ]);
  const valueIndex = findSettingsDraftColumn_(headers, [
    "value",
    "setting value",
    "current value",
    "configured value",
    "setting_value"
  ]);

  if (keyIndex === -1 || valueIndex === -1) {
    throw new Error("Settings draft must include Key and Value columns, or equivalent Setting and Current Value columns.");
  }

  const editableFields = getEditableSettingFields_();
  const imported = {};
  const ignoredKeys = [];

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][keyIndex] || "").trim();
    if (!key) continue;

    if (!editableFields[key]) {
      ignoredKeys.push(key);
      continue;
    }

    imported[key] = values[i][valueIndex];
  }

  const result = saveDashboardSettings(imported);

  return {
    ok: true,
    importedCount: Object.keys(imported).length,
    ignoredKeys,
    settings: result
  };
}

function findSettingsDraftColumn_(headers, acceptedNames) {
  for (const name of acceptedNames) {
    const index = headers.indexOf(name);
    if (index !== -1) return index;
  }

  return -1;
}

function getEditableSettingFields_() {
  const editableFields = {};

  getSettingSchema_().forEach(section => {
    section.fields.forEach(field => {
      if (!field.readonly) editableFields[field.key] = field;
    });
  });

  return editableFields;
}

function validateDashboardSettings_(values, schema) {
  const errors = [];
  const errorsByKey = {};

  function addError(field, message) {
    errors.push(`${field.label}: ${message}`);
    if (!errorsByKey[field.key]) errorsByKey[field.key] = [];
    errorsByKey[field.key].push(message);
  }

  schema.forEach(section => {
    section.fields.forEach(field => {
      if (field.readonly) return;

      const value = values[field.key];
      const text = String(value === null || value === undefined ? "" : value).trim();

      if (field.required && !text) {
        addError(field, "required.");
        return;
      }

      if (!text) return;

      if (field.type === "email" && !isValidEmail_(text)) {
        addError(field, "enter a valid email address.");
      }

      if (field.type === "url" && !isValidUrl_(text)) {
        addError(field, "enter a valid URL starting with http:// or https://.");
      }

      if (field.type === "color" && !/^#[0-9a-fA-F]{6}$/.test(text)) {
        addError(field, "use a hex colour like #F8F6F2.");
      }

      if (field.type === "number") {
        const n = Number(text);

        if (isNaN(n)) {
          addError(field, "enter a number.");
        } else if (field.min !== undefined && n < field.min) {
          addError(field, `must be at least ${field.min}.`);
        }
      }

      if (field.type === "date" && !/^\d{4}-\d{2}-\d{2}$/.test(text)) {
        addError(field, "use YYYY-MM-DD.");
      }

      if (field.type === "textarea" && field.key === "CALENDAR_ATTENDEES") {
        String(text)
          .split(/\r?\n|,/)
          .map(item => item.trim())
          .filter(Boolean)
          .forEach(email => {
            if (!isValidEmail_(email)) {
              addError(field, `${email} is not a valid email address.`);
            }
          });
      }
    });
  });

  return {
    ok: errors.length === 0,
    errors,
    errorsByKey
  };
}

function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function isValidUrl_(value) {
  return /^https?:\/\/\S+$/i.test(String(value || "").trim());
}

function testDashboardSettingsValidation() {
  const schema = getSettingSchema_();
  const defaults = {};
  const invalid = {};
  const editableFields = getEditableSettingFields_();

  schema.forEach(section => {
    section.fields.forEach(field => {
      defaults[field.key] = normaliseSettingValue_(field.fallback, field);
      invalid[field.key] = defaults[field.key];
    });
  });

  const defaultValidation = validateDashboardSettings_(defaults, schema);
  if (!defaultValidation.ok) {
    throw new Error("Default settings failed validation: " + defaultValidation.errors.join(" "));
  }

  invalid.APP_NAME = "";
  invalid.PRINTER_EMAIL = "not-an-email";
  invalid.CALENDAR_ATTENDEES = "good@example.com, bad-email";
  invalid.CALENDAR_EVENT_DURATION_MINUTES = "0";
  invalid.COLOUR_PRIMARY = "purple";
  invalid.FIKA_LOGO_URL = "not-a-url";

  const invalidValidation = validateDashboardSettings_(invalid, schema);
  const expectedKeys = [
    "APP_NAME",
    "PRINTER_EMAIL",
    "CALENDAR_ATTENDEES",
    "CALENDAR_EVENT_DURATION_MINUTES",
    "COLOUR_PRIMARY",
    "FIKA_LOGO_URL"
  ];

  expectedKeys.forEach(key => {
    if (!invalidValidation.errorsByKey[key]) {
      throw new Error("Expected validation error for " + key);
    }
  });

  if (!editableFields.APP_NAME || editableFields.LAST_INBOX_SCAN_AT) {
    throw new Error("Editable settings map did not match expected editable/read-only fields.");
  }

  return {
    ok: true,
    defaultErrorCount: defaultValidation.errors.length,
    invalidErrorCount: invalidValidation.errors.length
  };
}

function normaliseSettingValue_(value, field) {
  if (field.type === "number") {
    const n = Number(value);
    return isNaN(n) ? "" : n;
  }

  if (field.type === "textarea") {
    return String(value || "")
      .split(/\r?\n|,/)
      .map(item => item.trim())
      .filter(Boolean)
      .join(", ");
  }

  return String(value || "").trim();
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

function getConfiguredValue_(key, fallback) {
  const value = getSetting_(key, fallback);
  if (key === "COLOUR_PRIMARY" && String(value).trim().toUpperCase() === "#202020") {
    return fallback;
  }

  if (key === "COLOUR_SECONDARY" && String(value).trim().toUpperCase() === "#8A8A8A") {
    return fallback;
  }

  return value === "" || value === null || value === undefined ? fallback : value;
}

function getConfiguredNumber_(key, fallback) {
  const value = Number(getConfiguredValue_(key, fallback));
  return isNaN(value) ? fallback : value;
}
