function getWorkforceSettingsData() {
  const spreadsheet = getWorkforceSpreadsheet_();
  return {
    ok: true,
    managers: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.managers)),
    agencyContacts: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.agencyContacts)),
    roles: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.roleLibrary)),
    shiftPatterns: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.shiftPatterns)),
    sites: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.sites)),
    staff: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory))
      .filter(function(person) {
        return String(person.Name || "").trim() &&
          String(person["Employment Status"] || "").toLowerCase() !== "terminated" &&
          !workforceBoolean_(person.Terminated);
      })
      .map(function(person) {
        return {
          employeeId: person["Employee ID"],
          name: person.Name,
          role: person.Role,
          reliefTeam: workforceBoolean_(person["Relief Team"])
        };
      }),
    rotaTemplates: readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.rotaTemplates))
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

function saveWorkforceRotaTemplate(template) {
  const spreadsheet = getWorkforceSpreadsheet_();
  const siteId = String(template.siteId || template["Site ID"] || "").trim();
  const siteName = String(template.siteName || template["Site Name"] || "").trim() ||
    getWorkforceSiteName_(spreadsheet, siteId);
  const weekday = String(template.weekday || template.Weekday || "").trim();
  const role = String(template.role || template.Role || "").trim();
  const employeeName = String(template.employeeName || template["Employee Name"] || "").trim();
  const standardStatus = String(template.standardStatus || template["Standard Status"] || "IN").trim() || "IN";
  if (!siteId) throw new Error("Choose a site for this rota template.");
  if (!weekday) throw new Error("Choose a weekday for this rota template.");
  if (!role) throw new Error("Role is required.");
  if (!employeeName && standardStatus.toUpperCase() === "IN") {
    throw new Error("Employee name is required for an IN rota row.");
  }
  const id = [
    "template",
    siteId,
    weekday,
    slugifyWorkforce_(role),
    slugifyWorkforce_(employeeName || standardStatus)
  ].join("_");
  upsertWorkforceRows_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.rotaTemplates), "Template ID", [{
    "Template ID": id,
    "Site ID": siteId,
    "Site Name": siteName || siteId,
    "Weekday": weekday,
    "Role": role,
    "Employee Name": employeeName,
    "Standard Status": standardStatus,
    "Source": "Web App",
    "Observations": template.notes || "",
    "Active": true
  }]);
  return {
    ok: true,
    templateId: id,
    message: "Standard rota row saved."
  };
}

function getWorkforceSiteName_(spreadsheet, siteId) {
  const cleanSiteId = String(siteId || "");
  const site = readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.sites))
    .filter(function(row) {
      return String(row["Site ID"] || "") === cleanSiteId;
    })[0];
  return site ? String(site["Site Name"] || "") : "";
}

function isWorkforceEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}
