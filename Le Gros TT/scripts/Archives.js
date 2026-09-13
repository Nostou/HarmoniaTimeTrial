/********************
 * INIT
 ********************/

function initMaps(sheetName) {
  const sheet = getSheet(sheetName);  
  const mapLabels = getMapDefinitions().map(map => map.label);

  let output = [[""]];
  mapLabels.forEach(label => output.push([label]));
  output.push([""]);
  mapLabels.forEach(label => output.push([label]));

  sheet.getRange(1, 2, output.length, 1).setValues(output);
}

function initPlayers() {
  const archives = getSheet("Archives");
  const players = getArchivePlayers();

  archives
    .getRange(1, 3, 1, players.length)
    .setValues([players]);
}


/********************
 * RICH TEXT UTILS
 ********************/

function makePlainRichText(value) {
  return SpreadsheetApp
    .newRichTextValue()
    .setText(value == null ? "" : String(value))
    .build();
}

function setRichBlock(sheet, richGrid, valueGrid, startRow, rowCount, startCol, colCount) {
  const range = sheet.getRange(startRow + 1, startCol + 1, rowCount, colCount);
  const output = [];

  for (let r = 0; r < rowCount; r++) {
    const row = [];

    for (let c = 0; c < colCount; c++) {
      const richValue = richGrid[startRow + r][startCol + c];
      row.push(richValue || makePlainRichText(valueGrid[startRow + r][startCol + c]));
    }

    output.push(row);
  }

  range.setRichTextValues(output);
}


/********************
 * MAIN REFRESH
 ********************/

function refreshArchives() {
  //if (!assertOwnerOnly()) return;

  clearAllCaches();
  clearUILogs();

  initMaps("Archives");
  initPlayers();

  const tablesData = getSheetData("Tables");
  const maps = getMaps();
  const players = getArchivePlayers();
  const playerSet = new Set(players);

  const results = {};
  maps.forEach(m => {
    results[m] = {};
    players.forEach(p => {
      results[m][p] = { Shroom: 9999999, NI: 9999999 };
    });
  });

  const tableBlocks = getTableBlocks(tablesData, maps);

  tableBlocks.forEach(block => {
    const headerValues = tablesData[block.headerRow - 1] || [];

    block.mapRows.forEach(mapRow => {
      const row = tablesData[mapRow.row - 1] || [];
      const type = mapRow.format;
      const typeColumn = mapRow.formatColumn || block.formatColumn;
      if (!results[mapRow.map] || !type || typeColumn === null) return;

      for (let column = block.startColumn + 1; column < typeColumn; column++) {
        const newTime = TIME_TO_MS(row[column]);
        if (newTime === null) continue;

        const player = headerValues[column];
        if (!playerSet.has(player)) continue;

        if (newTime < results[mapRow.map][player][type]) {
          results[mapRow.map][player][type] = newTime;
        }
      }
    });
  });

  const archivesSheet = getSheet("Archives");
  const oldAverages = readAveragesFromSheet(archivesSheet, maps);

  computeAverageResults(results, maps, players);
  logAverageChanges(oldAverages, results, maps);

  const rankCache = buildRankCache();
  fillArchives(results, rankCache);

  showUILogs();
}


/********************
 * ARCHIVES RENDER
 ********************/

function fillArchives(results, rankCache) {
  const sheet = getSheet("Archives");

  const range = sheet.getDataRange();
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const maps = getMaps();
  const players = getArchivePlayers();

  const height = values.length;
  const width = values[0].length;

  const rich = Array.from({ length: height }, () =>
    Array.from({ length: width }, () => null)
  );

  function fillBlock(type, startIndex, isShroom) {
    for (let i = startIndex; i < startIndex + maps.length; i++) {
      const map = maps[i - startIndex];
      values[i][1] = getMapLabel_(map);
      if (!results[map]) continue;

      for (let j = 2; j < players.length + 2; j++) {
        const player = values[0][j];
        const v = results[map][player][type];

        if (v === 9999999) {
          values[i][j] = "-";
          rich[i][j] = null;
          backgrounds[i][j] = "#ffffff";
          continue;
        }

        const timeStr = MS_TO_TIME(v);
        const rankInfo = getRank(rankCache, map, timeStr, isShroom);
        const rankLabel = rankInfo ? rankInfo.rank : "Unranked";

        values[i][j] = "";
        rich[i][j] = buildTimeRankRichText(timeStr, rankLabel);
        backgrounds[i][j] = rankInfo ? rankInfo.color : "#ffffff";
      }
    }
  }

  fillBlock("Shroom", 1, true);
  fillBlock("NI", maps.length + 2, false);

  sheet
    .getRange(1, 1, values.length, values[0].length)
    .setValues(values);

  setRichBlock(sheet, rich, values, 1, maps.length, 2, players.length);
  setRichBlock(sheet, rich, values, maps.length + 2, maps.length, 2, players.length);

  sheet
    .getRange(1, 1, backgrounds.length, backgrounds[0].length)
    .setBackgrounds(backgrounds);
}
