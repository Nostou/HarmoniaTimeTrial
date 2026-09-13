function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const menu = ui.createMenu("Custom")
  .addItem("Create Table", "CreateTable")
  .addItem("Color selected Table", "colorSelectedTable")
  .addSeparator()
  .addItem("Refresh Archives", "refreshArchives")
  .addItem("Refresh Ranks", "refreshRanks")
  .addItem("Refresh Maps %", "openRefreshMapChances")
  .addItem("Refresh Stats", "Stats")
  .addToUi();
}
