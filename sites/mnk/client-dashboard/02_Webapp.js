function doGet(event) {
  const template = HtmlService.createTemplateFromFile("Index");
  template.portalAccessKey = String(event && event.parameter && event.parameter.key || "");
  return template.evaluate()
    .setTitle(CLIENT_PORTAL_CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
