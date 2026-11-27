/********************
 * INIT
 ********************/

function initMaps(sheetName) {
  const sheet = getSheet(sheetName);  
  const mapLabels = getMapDefinitions().map(map => map.label);

  const labels = mapLabels.map(label => [label]);
  sheet.getRange(2, 3, labels.length, 1).setValues(labels);
  sheet.getRange(labels.length + 3, 3, labels.length, 1).setValues(labels);
}

function initPlayers() {
  const archives = getSheet("Archives");
  const players = getArchivePlayers();

  archives
    .getRange(1, 4, 1, players.length)
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

  const maps = getMaps();
  const archivesSheet = getSheet("Archives");
  if (normalizeMapId_(archivesSheet.getRange(maps.length + 3, 3).getDisplayValue()) !== maps[0]) {
    throw new Error("The NI section in Archives does not match the map count. Adjust the sections and separator before refreshing.");
  }
  const oldAverages = readAveragesFromSheet(archivesSheet, maps);
  const rankCache = buildRankCache();
  initMaps("Archives");
  initPlayers();

  const tablesData = getSheetData("Tables");
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

  computeAverageResults(results, maps, players);
  logAverageChanges(oldAverages, results, maps);

  fillArchives(results, rankCache);

  showUILogs();
}


/********************
 * ARCHIVES RENDER
 ********************/

function fillArchives(results, rankCache) {
  const sheet = getSheet("Archives");
  const maps = getMaps();
  const players = getArchivePlayers();

  function fillBlock(type, startRow, isShroom) {
    const colors = [];
    const rich = maps.map(map => {
      const rowColors = [];
      const row = players.map(player => {
        const value = results[map][player][type];
        if (value === 9999999) {
          rowColors.push("#ffffff");
          return makePlainRichText("-");
        }
        const time = MS_TO_TIME(value);
        const rank = getRank(rankCache, map, time, isShroom);
        rowColors.push(rank ? rank.color : "#ffffff");
        return buildTimeRankRichText(time, rank ? rank.rank : "Unranked");
      });
      colors.push(rowColors);
      return row;
    });
    const range = sheet.getRange(startRow, 4, maps.length, players.length);
    range.setRichTextValues(rich);
    range.setBackgrounds(colors);
  }

  fillBlock("Shroom", 2, true);
  fillBlock("NI", maps.length + 3, false);
}
