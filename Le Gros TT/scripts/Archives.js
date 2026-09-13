/********************
 * INIT
 ********************/

function initMaps(sheetName) {
  const sheet = getSheet(sheetName);  
  const maps = getMaps();

  let output = [[""]];
  maps.forEach(m => output.push([m]));
  output.push([""]);
  maps.forEach(m => output.push([m]));

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

  let currentType = null;
  let typeIndex = 0;

  for (let i = 0; i < tablesData.length; i++) {
    const row = tablesData[i];
    const mapName = row[0];
    if (!results[mapName]) continue;

    for (let j = 2; j < row.length; j++) {
      if (row[j] === "") break;
      if (row[j] === "Shroom" || row[j] === "NI") {
        currentType = row[j];
        typeIndex = j;
        break;
      }
    }

    for (let j = 1; j < typeIndex; j++) {
      const newTime = TIME_TO_MS(row[j]);
      if (newTime === null) continue;

      for (let k = i; k >= 0; k--) {
        const player = tablesData[k][j];
        if (player === "") break;   
        if (!playerSet.has(player)) continue;        

        if (newTime < results[mapName][player][currentType]) {
          results[mapName][player][currentType] = newTime;
        }
        break;
      }
    }
  }

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

  const values = sheet.getDataRange().getValues();
  const backgrounds = sheet.getDataRange().getBackgrounds();
  const maps = getMaps();
  const players = getArchivePlayers();

  const height = values.length;
  const width = values[0].length;

  const rich = Array.from({ length: height }, () =>
    Array.from({ length: width }, () => null)
  );

  function fillBlock(type, startIndex, isShroom) {
    for (let i = startIndex; i < startIndex + maps.length; i++) {
      const map = values[i][1];

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
