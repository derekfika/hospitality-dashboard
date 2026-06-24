function setBrightHrApiCredentials(clientId, clientSecret, tokenUrl, apiBaseUrl) {
  const cleanClientId = String(clientId || "").trim();
  const cleanClientSecret = String(clientSecret || "").trim();
  const cleanTokenUrl = String(tokenUrl || WORKFORCE_CONFIG.brightHr.tokenUrl).trim();
  const cleanApiBaseUrl = String(apiBaseUrl || WORKFORCE_CONFIG.brightHr.apiBaseUrl).trim().replace(/\/+$/, "");

  if (!cleanClientId) throw new Error("BrightHR client ID is required.");
  if (!cleanClientSecret) throw new Error("BrightHR client secret is required.");
  if (cleanTokenUrl && !/^https:\/\/\S+$/i.test(cleanTokenUrl)) {
    throw new Error("BrightHR token URL must be an HTTPS URL.");
  }
  if (cleanApiBaseUrl && !/^https:\/\/\S+$/i.test(cleanApiBaseUrl)) {
    throw new Error("BrightHR API base URL must be an HTTPS URL.");
  }

  const keys = WORKFORCE_CONFIG.scriptProperties;
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    [keys.brightHrClientId]: cleanClientId,
    [keys.brightHrClientSecret]: cleanClientSecret
  }, false);
  properties.setProperty(keys.brightHrTokenUrl, cleanTokenUrl);
  properties.setProperty(keys.brightHrApiBaseUrl, cleanApiBaseUrl);
  clearBrightHrCachedToken_();

  return {
    ok: true,
    provider: WORKFORCE_CONFIG.hrProvider,
    clientIdStored: true,
    clientSecretStored: true,
    tokenUrlStored: true,
    apiBaseUrlStored: true,
    message: "BrightHR credentials stored in Script Properties. Secret value was not returned."
  };
}

function getBrightHrApiStatus() {
  const keys = WORKFORCE_CONFIG.scriptProperties;
  const properties = PropertiesService.getScriptProperties();
  return {
    ok: true,
    provider: WORKFORCE_CONFIG.hrProvider,
    hasClientId: Boolean(properties.getProperty(keys.brightHrClientId)),
    hasClientSecret: Boolean(properties.getProperty(keys.brightHrClientSecret)),
    hasTokenUrl: Boolean(properties.getProperty(keys.brightHrTokenUrl)),
    hasApiBaseUrl: Boolean(properties.getProperty(keys.brightHrApiBaseUrl)),
    hasAbsencesPath: Boolean(properties.getProperty(keys.brightHrAbsencesPath)),
    absencesPath: properties.getProperty(keys.brightHrAbsencesPath) || "",
    absencesMethod: properties.getProperty(keys.brightHrAbsencesMethod) || "",
    hasCachedAccessToken: Boolean(properties.getProperty(keys.brightHrAccessToken)),
    cachedTokenExpiresAt: properties.getProperty(keys.brightHrAccessTokenExpiresAt) || ""
  };
}

function clearBrightHrApiCredentials() {
  const keys = WORKFORCE_CONFIG.scriptProperties;
  PropertiesService.getScriptProperties().deleteAllProperties();
  return {
    ok: true,
    message: "Workforce Script Properties cleared for this Apps Script project."
  };
}

function testBrightHrConnection() {
  const token = getBrightHrAccessToken_();
  return {
    ok: true,
    provider: WORKFORCE_CONFIG.hrProvider,
    accessTokenReceived: Boolean(token),
    message: "BrightHR token request succeeded."
  };
}

function testBrightHrEmployeesQuery(limit) {
  const data = brightHrRequest_("employees/v1/query", {
    method: "post",
    payload: {
      pageSize: Math.max(1, Math.min(Number(limit || 5), 50))
    }
  });
  return {
    ok: true,
    provider: WORKFORCE_CONFIG.hrProvider,
    endpoint: "/employees/v1/query",
    resultType: Object.prototype.toString.call(data),
    sample: data
  };
}

function syncBrightHrEmployees() {
  setupWorkforceOperationsPlatform();
  const spreadsheet = getWorkforceSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory);
  const existingStaff = buildExistingStaffIndex_(sheet);
  const employees = fetchBrightHrEmployees_();
  const now = new Date();
  const rows = employees.map(function(employee) {
    return normaliseBrightHrEmployee_(employee, now, existingStaff);
  });
  upsertWorkforceRows_(sheet, "Employee ID", rows);
  return {
    ok: true,
    provider: WORKFORCE_CONFIG.hrProvider,
    sheet: WORKFORCE_CONFIG.sheets.staffDirectory,
    synced: rows.length
  };
}

function setBrightHrAbsencesEndpoint(path, method) {
  const cleanPath = String(path || "").trim().replace(/^\/+/, "");
  const cleanMethod = String(method || "post").trim().toLowerCase();
  if (!cleanPath) throw new Error("BrightHR absence endpoint path is required.");
  if (["get", "post"].indexOf(cleanMethod) === -1) throw new Error("BrightHR absence method must be GET or POST.");
  PropertiesService.getScriptProperties()
    .setProperty(WORKFORCE_CONFIG.scriptProperties.brightHrAbsencesPath, cleanPath);
  PropertiesService.getScriptProperties()
    .setProperty(WORKFORCE_CONFIG.scriptProperties.brightHrAbsencesMethod, cleanMethod);
  return {
    ok: true,
    setting: WORKFORCE_CONFIG.scriptProperties.brightHrAbsencesPath,
    path: cleanPath,
    method: cleanMethod
  };
}

function syncBrightHrAbsences() {
  setupWorkforceOperationsPlatform();
  const keys = WORKFORCE_CONFIG.scriptProperties;
  const properties = PropertiesService.getScriptProperties();
  let path = properties.getProperty(keys.brightHrAbsencesPath);
  let method = properties.getProperty(keys.brightHrAbsencesMethod) || "post";
  if (!path) {
    const discovered = discoverBrightHrAbsenceEndpoint_();
    path = discovered.path;
    method = discovered.method;
    properties.setProperty(keys.brightHrAbsencesPath, path);
    properties.setProperty(keys.brightHrAbsencesMethod, method);
  }

  const spreadsheet = getWorkforceSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.absences);
  const staffIndex = buildStaffLookupForAbsences_(spreadsheet);
  const items = fetchBrightHrAbsences_(path, method);
  const now = new Date();
  const rows = items.map(function(absence, index) {
    return normaliseBrightHrAbsence_(absence, now, index, staffIndex);
  }).filter(function(row) {
    return row["Employee Name"] && row["Start Date"];
  });
  upsertWorkforceRows_(sheet, "Absence ID", rows);
  return {
    ok: true,
    provider: WORKFORCE_CONFIG.hrProvider,
    endpoint: path,
    method: method,
    sheet: WORKFORCE_CONFIG.sheets.absences,
    synced: rows.length,
    message: rows.length + " absence row(s) synced from BrightHR."
  };
}

function discoverBrightHrAbsenceEndpoint() {
  const result = discoverBrightHrAbsenceEndpoint_();
  return {
    ok: true,
    path: result.path,
    method: result.method,
    itemCount: result.itemCount,
    message: "BrightHR absence endpoint found: " + result.method.toUpperCase() + " /" + result.path
  };
}

function fetchBrightHrAbsences_(path, method) {
  const cleanMethod = String(method || "post").toLowerCase();
  const items = [];
  let continuationToken = "";
  do {
    const options = { method: cleanMethod };
    if (cleanMethod === "post") {
      options.payload = { pageSize: 100 };
      if (continuationToken) options.payload.continuationToken = continuationToken;
    }
    const data = brightHrRequest_(path, options);
    items.push.apply(items, getBrightHrItems_(data));
    continuationToken = cleanMethod === "post" ? String(data.continuationToken || "") : "";
  } while (continuationToken);
  return items;
}

function discoverBrightHrAbsenceEndpoint_() {
  const candidates = [
    { path: "absences/v1/query", method: "post" },
    { path: "absence/v1/query", method: "post" },
    { path: "leave/v1/query", method: "post" },
    { path: "leaves/v1/query", method: "post" },
    { path: "holiday/v1/query", method: "post" },
    { path: "holidays/v1/query", method: "post" },
    { path: "sickness/v1/query", method: "post" },
    { path: "timeoff/v1/query", method: "post" },
    { path: "time-off/v1/query", method: "post" },
    { path: "employee-absences/v1/query", method: "post" }
  ];
  const errors = [];
  for (let index = 0; index < candidates.length; index++) {
    const candidate = candidates[index];
    try {
      const data = brightHrRequest_(candidate.path, {
        method: candidate.method,
        payload: { pageSize: 1 }
      });
      const items = getBrightHrItems_(data);
      if (isBrightHrCollectionResponse_(data)) {
        return {
          path: candidate.path,
          method: candidate.method,
          itemCount: items.length
        };
      }
    } catch (error) {
      errors.push(candidate.path + ": " + (error.message || String(error)).slice(0, 160));
    }
  }
  throw new Error(
    "Could not discover BrightHR absence endpoint. Ask BrightHR for the holiday/absence/sickness query path, then run setBrightHrAbsencesEndpoint(\"PATH\", \"post\"). Tried: " +
    errors.join(" | ")
  );
}

function isBrightHrCollectionResponse_(data) {
  return Array.isArray(data) ||
    Array.isArray(data.items) ||
    Array.isArray(data.results) ||
    Array.isArray(data.data);
}

function brightHrGet_(path) {
  return brightHrRequest_(path, { method: "get" });
}

function brightHrRequest_(path, options) {
  const keys = WORKFORCE_CONFIG.scriptProperties;
  const properties = PropertiesService.getScriptProperties();
  const apiBaseUrl = String(properties.getProperty(keys.brightHrApiBaseUrl) || "").trim();
  if (!apiBaseUrl) throw new Error("BRIGHTHR_API_BASE_URL is not set.");
  const cleanPath = String(path || "").replace(/^\/+/, "");
  if (!cleanPath) throw new Error("BrightHR API path is required.");
  const requestOptions = options || {};
  const fetchOptions = {
    method: requestOptions.method || "get",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + getBrightHrAccessToken_(),
      Accept: "application/json"
    }
  };
  if (requestOptions.payload !== undefined) {
    fetchOptions.contentType = "application/json";
    fetchOptions.payload = JSON.stringify(requestOptions.payload);
  }
  const response = UrlFetchApp.fetch(apiBaseUrl.replace(/\/+$/, "") + "/" + cleanPath, fetchOptions);
  return parseBrightHrResponse_(response);
}

function getBrightHrAccessToken_() {
  const keys = WORKFORCE_CONFIG.scriptProperties;
  const properties = PropertiesService.getScriptProperties();
  const cachedToken = properties.getProperty(keys.brightHrAccessToken);
  const expiresAt = Number(properties.getProperty(keys.brightHrAccessTokenExpiresAt) || 0);
  if (cachedToken && expiresAt && Date.now() < expiresAt - 60000) return cachedToken;

  const clientId = properties.getProperty(keys.brightHrClientId);
  const clientSecret = properties.getProperty(keys.brightHrClientSecret);
  const tokenUrl = properties.getProperty(keys.brightHrTokenUrl);
  if (!clientId || !clientSecret) {
    throw new Error("BrightHR client ID/secret are not set. Run setBrightHrApiCredentials().");
  }
  if (!tokenUrl) {
    throw new Error("BRIGHTHR_TOKEN_URL is not set. BrightHR needs to provide the OAuth token URL.");
  }

  const response = UrlFetchApp.fetch(tokenUrl, {
    method: "post",
    muteHttpExceptions: true,
    contentType: "application/x-www-form-urlencoded",
    payload: {
      grant_type: "client_credentials",
      client_id: clientId,
      client_secret: clientSecret
    },
    headers: {
      Accept: "application/json"
    }
  });
  const data = parseBrightHrResponse_(response);
  const accessToken = data.access_token || data.accessToken || "";
  if (!accessToken) {
    throw new Error("BrightHR token response did not include an access token.");
  }
  const expiresInSeconds = Number(data.expires_in || data.expiresIn || 3600);
  properties.setProperty(keys.brightHrAccessToken, accessToken);
  properties.setProperty(
    keys.brightHrAccessTokenExpiresAt,
    String(Date.now() + expiresInSeconds * 1000)
  );
  return accessToken;
}

function parseBrightHrResponse_(response) {
  const status = response.getResponseCode();
  const text = response.getContentText();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch (error) {
    data = { raw: text };
  }
  if (status < 200 || status >= 300) {
    throw new Error(
      "BrightHR API returned HTTP " + status + ": " +
      String(text || "").slice(0, 500)
    );
  }
  return data;
}

function clearBrightHrCachedToken_() {
  const keys = WORKFORCE_CONFIG.scriptProperties;
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(keys.brightHrAccessToken);
  properties.deleteProperty(keys.brightHrAccessTokenExpiresAt);
}

function fetchBrightHrEmployees_() {
  let continuationToken = "";
  const employees = [];
  do {
    const payload = { pageSize: 100 };
    if (continuationToken) payload.continuationToken = continuationToken;
    const data = brightHrRequest_("employees/v1/query", {
      method: "post",
      payload: payload
    });
    employees.push.apply(employees, getBrightHrItems_(data));
    continuationToken = data.continuationToken || "";
  } while (continuationToken);
  return employees;
}

function getBrightHrItems_(data) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (Array.isArray(data.items)) return data.items;
  if (Array.isArray(data.results)) return data.results;
  if (Array.isArray(data.data)) return data.data;
  return [];
}

function normaliseBrightHrEmployee_(employee, syncedAt, existingStaff) {
  const name = employee.name || {};
  const employment = employee.employment || {};
  const metadata = employee._metadata || {};
  const fullName = [
    name.givenName || name.firstName || "",
    name.familyName || name.lastName || ""
  ].filter(Boolean).join(" ") || employee.fullName || employee.displayName || "";
  const terminated = Boolean(metadata.isTerminated || employment.end);
  const existing = findExistingBrightHrStaff_(existingStaff, employee.id || employee.employeeId || "", fullName);
  return {
    "Employee ID": employee.id || employee.employeeId || "",
    "Name": fullName,
    "Email": employee.email || "",
    "External Reference": employee.externalReference || "",
    "Role": employment.jobTitle || employee.jobTitle || "",
    "Primary Site": existing["Primary Site"] || "",
    "Secondary Sites": existing["Secondary Sites"] || "",
    "Contract Hours": existing["Contract Hours"] || "",
    "Employment Status": terminated ? "Terminated" : "Active",
    "Registered": metadata.isRegistered === undefined ? "" : Boolean(metadata.isRegistered),
    "Terminated": terminated,
    "Relief Team": existing["Relief Team"] || false,
    "Event Team": existing["Event Team"] || false,
    "Coffee Trainer": existing["Coffee Trainer"] || false,
    "Manager": existing.Manager || false,
    "BrightHR Raw JSON": JSON.stringify(employee),
    "Last Synced": syncedAt
  };
}

function normaliseBrightHrAbsence_(absence, syncedAt, index, staffIndex) {
  const employee = absence.employee || absence.employeeDetails || {};
  const employeeName = typeof employee.name === "object" ? employee.name : {};
  const employeeId = absence.employeeId || absence.employeeID || employee.id || employee.employeeId || "";
  const fullName = [
    employeeName.givenName || employeeName.firstName || "",
    employeeName.familyName || employeeName.lastName || ""
  ].filter(Boolean).join(" ") ||
    (typeof employee.name === "string" ? employee.name : "") ||
    employee.fullName ||
    employee.displayName ||
    absence.employeeName ||
    getStaffNameForAbsence_(staffIndex, employeeId) ||
    "";
  const startDate = absence.start || absence.startDate || absence.dateFrom || absence.from || absence.periodStart || (absence.period && absence.period.start) || "";
  const endDate = absence.end || absence.endDate || absence.dateTo || absence.to || absence.periodEnd || (absence.period && absence.period.end) || startDate;
  const absenceType = absence.type || absence.absenceType || absence.category || absence.reason || absence.leaveType || absence.name || "";
  const id = absence.id || absence.absenceId || [
    "brighthr_absence",
    employeeId || slugifyWorkforce_(fullName),
    normaliseBrightHrDateForId_(startDate),
    slugifyWorkforce_(absenceType || index)
  ].join("_");
  return {
    "Absence ID": id,
    "Employee ID": employeeId,
    "Employee Name": fullName,
    "Absence Type": absenceType,
    "Start Date": startDate,
    "End Date": endDate,
    "Status": absence.status || absence.approvalStatus || "",
    "Source": "BrightHR",
    "BrightHR Raw JSON": JSON.stringify(absence),
    "Last Synced": syncedAt
  };
}

function buildExistingStaffIndex_(sheet) {
  const index = { byId: {}, byName: {} };
  readWorkforceObjects_(sheet).forEach(function(row) {
    const id = String(row["Employee ID"] || "").trim();
    const name = normaliseWorkforcePerson_(row.Name);
    if (id) index.byId[id] = row;
    if (name) index.byName[name] = row;
  });
  return index;
}

function findExistingBrightHrStaff_(existingStaff, employeeId, name) {
  const index = existingStaff || { byId: {}, byName: {} };
  return index.byId[String(employeeId || "").trim()] ||
    index.byName[normaliseWorkforcePerson_(name)] ||
    {};
}

function buildStaffLookupForAbsences_(spreadsheet) {
  const lookup = { byId: {}, byName: {} };
  readWorkforceObjects_(spreadsheet.getSheetByName(WORKFORCE_CONFIG.sheets.staffDirectory))
    .forEach(function(person) {
      const id = String(person["Employee ID"] || "").trim();
      const name = String(person.Name || "").trim();
      if (id) lookup.byId[id] = name;
      if (name) lookup.byName[normaliseWorkforcePerson_(name)] = name;
    });
  return lookup;
}

function getStaffNameForAbsence_(staffIndex, employeeId) {
  return staffIndex && staffIndex.byId
    ? staffIndex.byId[String(employeeId || "").trim()] || ""
    : "";
}

function normaliseBrightHrDateForId_(value) {
  const text = String(value || "").slice(0, 10);
  return text ? text.replace(/[^0-9]/g, "") : "unknown_date";
}

function upsertWorkforceRows_(sheet, keyHeader, rowObjects) {
  if (!sheet) throw new Error("Target workforce sheet was not found.");
  if (!rowObjects.length) return;
  const map = workforceHeaderMap_(sheet);
  const keyColumn = map[keyHeader];
  if (!keyColumn) throw new Error("Sheet is missing key column: " + keyHeader);
  const existing = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, keyColumn, sheet.getLastRow() - 1, 1)
      .getDisplayValues()
      .forEach(function(row, index) {
        const key = String(row[0] || "").trim();
        if (key) existing[key] = index + 2;
      });
  }
  rowObjects.forEach(function(rowObject) {
    const key = String(rowObject[keyHeader] || "").trim();
    if (!key) return;
    const values = Object.keys(map).map(function(header) {
      return Object.prototype.hasOwnProperty.call(rowObject, header)
        ? rowObject[header]
        : "";
    });
    if (existing[key]) {
      sheet.getRange(existing[key], 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
  });
}

function workforceHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  const map = {};
  headers.forEach(function(header, index) {
    map[String(header).trim()] = index + 1;
  });
  return map;
}
