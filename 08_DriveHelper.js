function convertXlsxToGoogleSheet_(xlsxBlob) {
  const tempXlsx = DriveApp.createFile(xlsxBlob);
  tempXlsx.setName(`TEMP_BOOKING_${Date.now()}.xlsx`);

  const resource = {
    title: tempXlsx.getName().replace(/\.xlsx$/i, ""),
    mimeType: MimeType.GOOGLE_SHEETS
  };

  const converted = Drive.Files.copy(resource, tempXlsx.getId(), { convert: true });
  tempXlsx.setTrashed(true);

  return converted.id;
}

function authorizeDocsAccess() {
  const doc = DocumentApp.create("TEMP_DOC_AUTH_TEST");
  const id = doc.getId();
  doc.saveAndClose();
  DriveApp.getFileById(id).setTrashed(true);

  SpreadsheetApp.getUi().alert("Docs access authorised.");
}