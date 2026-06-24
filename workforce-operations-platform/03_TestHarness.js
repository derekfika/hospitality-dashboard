function runWorkforcePlatformTests() {
  const status = getBrightHrApiStatus();
  return {
    ok: true,
    tests: [
      {
        name: "BrightHR status can be read without exposing secrets",
        ok: status.ok === true &&
          Object.prototype.hasOwnProperty.call(status, "hasClientSecret")
      },
      {
        name: "Workforce config has core rota sheets",
        ok: Boolean(WORKFORCE_CONFIG.sheets.staffDirectory) &&
          Boolean(WORKFORCE_CONFIG.sheets.rotaShifts)
      }
    ]
  };
}

