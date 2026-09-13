let _mapsCache = null;
let _playersCache = null;

function clearAllCaches() {
  _mapsCache = null;
  _playersCache = null;
}

function getMaps() {
  if (_mapsCache) return _mapsCache;
  _mapsCache = getDropdownValues("Template", "A2");
  return _mapsCache;
}

function getPlayers() {
  if (_playersCache) return _playersCache;
  _playersCache = getDropdownValues("Template", "B1");
  return _playersCache;
}

function getArchivePlayers() {
  return ["Average", ...getPlayers()];
}

function getDropdownValues(sheetName, inCell) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  const cell = sheet.getRange(inCell);
  const rule = cell.getDataValidation();  

  const values = rule.getCriteriaValues()[0];
  if (!Array.isArray(values)) return [];

  return values.filter(v => v !== "-");
}

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(sheetName);
}

function getSheetData(sheetName) {
  return getSheet(sheetName).getDataRange().getValues();
}

function buildRankCache() {
  const sheet = getSheet("Ranks");
  const data = sheet.getDataRange().getValues();
  const backgrounds = sheet.getDataRange().getBackgrounds();
  const maps = getMaps();

  const cache = {
    Shroom: {},
    NI: {}
  };

  function fillCache(isShroom) {
    const startIndex = isShroom ? 1 : maps.length + 2;
    const type = isShroom ? "Shroom" : "NI";

    for (let i = startIndex; i < startIndex + maps.length; i++) {
      const mapName = data[i][1];
      cache[type][mapName] = [];

      for (let j = 3; j < 18; j++) {
        cache[type][mapName].push({
          rank: data[0][j],
          time: TIME_TO_MS(data[i][j]),
          color: backgrounds[i][j]
        });
      }
    }
  }

  fillCache(true);
  fillCache(false);

  return cache;
}

function buildTimeRankRichText(time, rank) {
  const text = `${time}\n${rank}`;

  const builder = SpreadsheetApp.newRichTextValue()
    .setText(text);

  // Style temps
  builder.setTextStyle(0, time.length, SpreadsheetApp.newTextStyle()
    .setFontSize(11)
    .setBold(true)
    .build()
  );

  // Style rank
  const rankStart = time.length + 1;
  builder.setTextStyle(rankStart, text.length, SpreadsheetApp.newTextStyle()
    .setFontSize(9)
    .build()
  );

  return builder.build();
}

function getRank(cache, mapName, time, isShroom) {
  const type = isShroom ? "Shroom" : "NI";
  const playerTime = TIME_TO_MS(time);

  const ranks = cache[type][mapName];
  if (!ranks) return null;

  for (const r of ranks) {
    if (r.time > playerTime) {
      return { rank: r.rank, color: r.color };
    }
  }

  return null;
}

function getRankMultipliers() {
  const sheet = getSheet("Ranks");
  const data = sheet.getRange("T2:U17").getValues();

  const multipliers = {};
  data.forEach(([rank, value]) => {
    if (rank && value) {
      multipliers[rank] = Number(value);
    }
  });

  return multipliers;
}