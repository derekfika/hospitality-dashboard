function runFeedbackPortalTests() {
  const tests = [];
  function record(name, condition) {
    tests.push({ name: name, ok: Boolean(condition) });
  }
  record("Spreadsheet URL extraction",
    extractFeedbackSpreadsheetId_("https://docs.google.com/spreadsheets/d/1ExampleSpreadsheetId123456789/edit") ===
    "1ExampleSpreadsheetId123456789");
  record("Rating validation", feedbackRating_(5) === 5 && feedbackRating_(6) === 0);
  record("NPS validation", feedbackNps_(0) === 0 && feedbackNps_(10) === 10);
  record("Item normalisation", normaliseFeedbackItems_([
    { itemId: "pastries", name: "Mini Pastries", qty: 2 }
  ])[0].quantity === 2);
  record("Due date calculation", Boolean(getFeedbackRequestDueDate_({
    eventDate: "2030-06-22",
    serviceTimes: ["09:00"],
    clientBooking: { event: { endTime: "10:00" } }
  })));
  record("Confirmed status is the only eligible status",
    FEEDBACK_SITES.angel_court.eligibleStatuses.length === 1 &&
    FEEDBACK_SITES.angel_court.eligibleStatuses[0] === "CONFIRMED");
  record("Angel Court registry entry",
    FEEDBACK_SITES.angel_court.siteId === "angel_court" &&
    FEEDBACK_SITES.angel_court.spreadsheetPropertyKey ===
      "ANGEL_COURT_DASHBOARD_SPREADSHEET_ID");
  return {
    ok: tests.every(function(test) { return test.ok; }),
    tests: tests
  };
}
