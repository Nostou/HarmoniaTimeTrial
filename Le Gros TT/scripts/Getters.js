let _mapsCache = null;
let _mapDefinitionsCache = null;
let _playersCache = null;
let _tablesCache = null;

const MAP_CATALOG_START_ROW = 7;
const MAP_CATALOG_START_COLUMN = 1;
const MAP_CATALOG_WIDTH = 6;
const TABLE_TEMPLATE_PLAYER_CELL = "I1";

function clearAllCaches() {
  _mapsCache = null;
  _mapDefinitionsCache = null;
  _playersCache = null;
  _tablesCache = null;
}

function getMaps() {
  if (_mapsCache) return _mapsCache;
  _mapsCache = getMapDefinitions().map(map => map.id);
  return _mapsCache;
}

function getMapDefinitions() {
  if (_mapDefinitionsCache) return _mapDefinitionsCache;

  const sheet = getSheet("Template");
  const lastRow = sheet.getLastRow();
  if (lastRow < MAP_CATALOG_START_ROW) return [];

  const values = sheet
    .getRange(
      MAP_CATALOG_START_ROW,
      MAP_CATALOG_START_COLUMN,
      lastRow - MAP_CATALOG_START_ROW + 1,
      MAP_CATALOG_WIDTH
    )
    .getDisplayValues();

  const seen = new Set();
  _mapDefinitionsCache = values
    .map(row => {
      const id = normalizeMapId_(row[0]);
      const tag = String(row[1] || "").trim();
      const name = String(row[2] || "").trim();
      if (!id) return null;
      if (seen.has(id)) throw new Error(`Duplicate map ID: ${id}.`);

      seen.add(id);
      return {
        id,
        tag,
        name,
        label: `${id} - ${name}`,
        count: Number(row[3]) || 0,
        weight: row[4] === "" ? null : Number(row[4]) || 0,
        chance: row[5]
      };
    })
    .filter(Boolean);

  return _mapDefinitionsCache;
}

function getPlayers() {
  if (_playersCache) return _playersCache;
  _playersCache = getDropdownValues("Template", TABLE_TEMPLATE_PLAYER_CELL);
  return _playersCache;
}

function normalizeMapId_(value) {
  const text = String(value === null || value === undefined ? "" : value).trim();
  const match = text.match(/^(\d{1,2})(?:\s*-\s*|$)/);
  if (!match) return text;
  return match[1].padStart(2, "0");
}

function getMapLabel_(mapId) {
  const normalizedId = normalizeMapId_(mapId);
  const definition = getMapDefinitions().find(map => map.id === normalizedId);
  return definition ? definition.label : String(mapId || "").trim();
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

  return [...new Set(values
    .map(value => String(value).trim())
    .filter(value => value && value !== "-"))];
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
  const range = sheet.getDataRange();
  const data = range.getValues();
  const backgrounds = range.getBackgrounds();
  const maps = getMaps();

  for (let i = 0; i < maps.length; i++) {
    if (normalizeMapId_((data[i + 1] || [])[2]) !== maps[i] ||
        normalizeMapId_((data[i + maps.length + 2] || [])[2]) !== maps[i]) {
      throw new Error("The map list in Ranks is out of date. Run Custom > Refresh Ranks first.");
    }
  }

  const cache = {
    Shroom: {},
    NI: {}
  };

  function fillCache(isShroom) {
    const startIndex = isShroom ? 1 : maps.length + 2;
    const type = isShroom ? "Shroom" : "NI";

    for (let i = startIndex; i < startIndex + maps.length; i++) {
      const mapId = normalizeMapId_(data[i][2]);
      cache[type][mapId] = [];

      for (let j = 4; j < 19; j++) {
        cache[type][mapId].push({
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

  const ranks = cache[type][normalizeMapId_(mapName)];
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
  const data = sheet.getRange("U2:V17").getValues();

  const multipliers = {};
  data.forEach(([rank, value]) => {
    if (rank && value) {
      multipliers[rank] = Number(value);
    }
  });

  return multipliers;
}
