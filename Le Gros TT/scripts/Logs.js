let _uiLogs = [];

function logUI(msg) {
  _uiLogs.push(String(msg));
}

function clearUILogs() {
  _uiLogs = [];
}

function showUILogs(title = "Logs") {
  if (_uiLogs.length === 0) return;

  SpreadsheetApp.getUi().alert(
    title,
    _uiLogs.join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}