function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const menu = ui.createMenu("Custom")
  .addItem("Create Shroom Table", "insertTemplateShroom")
  .addItem("Create NI Table", "insertTemplateNI")
  .addItem("Create Tracks Markdown","TracksMarkdown")
  .addItem("Color selected Table", "colorSelectedTable")
  .addSeparator()
  .addItem("Refresh Archives (Full)", "refreshArchives")
  .addItem("Refresh Ranks", "refreshRanks")
  .addItem("Refresh Maps %", "refreshMapChances")
  .addItem("Refresh Stats", "Stats")
  .addToUi();
}