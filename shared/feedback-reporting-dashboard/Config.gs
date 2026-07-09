const FEEDBACK_REPORTING_CONFIG = Object.freeze({
  appName: "Hospitality Feedback Reporting",
  timeZone: "Europe/London",
  sheets: {
    requests: "Feedback Requests",
    responses: "Feedback Responses",
    itemRatings: "Feedback Item Ratings"
  },
  colours: {
    ink: "#243036",
    muted: "#667176",
    paper: "#f4f3ef",
    surface: "#ffffff",
    primary: "#4f5d64",
    accent: "#75efb8",
    warning: "#8a6e3e",
    danger: "#8c4f4f"
  }
});

const FEEDBACK_REPORTING_SITES = Object.freeze([
  {
    siteId: "all",
    siteName: "All Sites",
    spreadsheetId: "",
    propertyKey: ""
  },
  {
    siteId: "demo",
    siteName: "Demo",
    spreadsheetId: "1YwoWOifOYIT35aZbAdxGxXVGRsrvRBPuJkPQNkuj3Ow",
    propertyKey: "DEMO_DASHBOARD_SPREADSHEET_ID"
  },
  {
    siteId: "mnk",
    siteName: "MNK",
    spreadsheetId: "1GIGIh_oAY0yLrrlXPaSvHte2oMPT8S_dKFAYdZN6nuc",
    propertyKey: "MNK_DASHBOARD_SPREADSHEET_ID"
  },
  {
    siteId: "angel_court",
    siteName: "Angel Court",
    spreadsheetId: "",
    propertyKey: "ANGEL_COURT_DASHBOARD_SPREADSHEET_ID"
  },
  {
    siteId: "cfc",
    siteName: "CFC",
    spreadsheetId: "",
    propertyKey: "CFC_DASHBOARD_SPREADSHEET_ID"
  },
  {
    siteId: "58ve",
    siteName: "58 Victoria Embankment",
    spreadsheetId: "",
    propertyKey: "VE58_DASHBOARD_SPREADSHEET_ID"
  }
]);
