function getWorkforceSettingsData() {
  const spreadsheet = getWorkforceSpreadsheet_();
  return {
    ok: true,
    managers: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.managers)),
    agencyContacts: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.agencyContacts)),
    roles: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.roleLibrary)),
    shiftPatterns: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.shiftPatterns))
  };
}

function saveWorkforceManager(manager) {
  const spreadsheet = getWorkforceSpreadsheet_();
  const name = String(manager.managerName || manager["Manager Name"] || "").trim();
  const email = String(manager.email || manager.Email || "").trim();
  if (!name) throw new Error("Manager name is required.");
  if (!isWorkforceEmail_(email)) throw new Error("Enter a valid manager email.");
  const id = "manager_" + slugifyWorkforce_(name);
  upsertWorkforceRows_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.managers), "Manager ID", [{
    "Manager ID": id,
    "Manager Name": name,
    "Email": email,
    "Phone": manager.phone || "",
    "Primary Site ID": manager.primarySiteId || "",
    "Secondary Site IDs": manager.secondarySiteIds || "",
    "Receives Gap Alerts": true,
    "Receives Agency Confirmations": true,
    "Active": true,
    "Notes": manager.notes || ""
  }]);
  return { ok: true, managerId: id, message: "Manager saved." };
}

function saveWorkforceAgencyContact(contact) {
  const spreadsheet = getWorkforceSpreadsheet_();
  const agencyName = String(contact.agencyName || contact["Agency Name"] || "").trim();
  const email = String(contact.email || contact.Email || "").trim();
  if (!agencyName) throw new Error("Agency name is required.");
  if (!isWorkforceEmail_(email)) throw new Error("Enter a valid agency email.");
  const id = "agency_" + slugifyWorkforce_(agencyName);
  upsertWorkforceRows_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.agencyContacts), "Agency ID", [{
    "Agency ID": id,
    "Agency Name": agencyName,
    "Contact Name": contact.contactName || "",
    "Email": email,
    "Phone": contact.phone || "",
    "Roles Supplied": contact.rolesSupplied || "",
    "Sites Covered": contact.sitesCovered || "",
    "Default Rate": contact.defaultRate || "",
    "Active": true,
    "Notes": contact.notes || ""
  }]);
  return { ok: true, agencyId: id, message: "Agency contact saved." };
}

function isWorkforceEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}
