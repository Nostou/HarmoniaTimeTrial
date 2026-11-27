/***********************
 * TABLE CREATION
 ***********************/

const CREATE_TABLE_DEFAULT_COUNT = 5;
const CREATE_TABLE_MAX_COUNT = 40;
const CREATE_TABLE_DEFAULT_PLAYER_COUNT = 5;
const CREATE_TABLE_MIN_PLAYER_COUNT = 2;
const CREATE_TABLE_FORMATS = ["Shroom", "NI"];
// Change this value when the oldest tables should no longer affect weighting.
const DEFAULT_IGNORED_TABLES_FOR_MAP_CHANCES = 59;

const TABLE_TEMPLATE_HEADER_ROW = 1;
const TABLE_TEMPLATE_SHROOM_ROW = 2;
const TABLE_TEMPLATE_NI_ROW = 3;
const TABLE_TEMPLATE_TOTAL_ROW = 4;
const TABLE_TEMPLATE_START_COLUMN = 8;
const TABLE_TEMPLATE_WIDTH = 3;

function CreateTable() {
  const html = HtmlService.createHtmlOutputFromFile("ui/CreateTablePopup")
    .setWidth(720)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, "Create Table");
}

// Kept for compatibility with existing manual shortcuts or triggers.
function insertTemplateShroom() {
  CreateTable();
}

// Kept for compatibility with existing manual shortcuts or triggers.
function insertTemplateNI() {
  CreateTable();
}

function getCreateTableOptions(mode, count, playerCount) {
  const normalizedMode = normalizeCreateTableMode_(mode);
  clearAllCaches();
  const mapDefinitions = getMapDefinitions();
  const maps = mapDefinitions.map(map => map.id);
  const availablePlayers = getPlayers();
  if (maps.length === 0) throw new Error("No maps are available in Template.");
  if (availablePlayers.length < CREATE_TABLE_MIN_PLAYER_COUNT) {
    throw new Error("At least 2 players must be available in Template.");
  }

  const requestedCount = normalizedMode === "Custom"
    ? normalizeCreateTableCount_(count, maps.length)
    : CREATE_TABLE_DEFAULT_COUNT;

  if (requestedCount > maps.length) {
    throw new Error(`At least ${requestedCount} maps must be available.`);
  }

  const requestedPlayerCount = normalizePlayerCount_(playerCount, availablePlayers.length);

  return {
    mode: normalizedMode,
    count: requestedCount,
    playerCount: requestedPlayerCount,
    defaultCount: CREATE_TABLE_DEFAULT_COUNT,
    defaultPlayerCount: Math.min(CREATE_TABLE_DEFAULT_PLAYER_COUNT, availablePlayers.length),
    minCount: 1,
    maxCount: Math.min(CREATE_TABLE_MAX_COUNT, maps.length),
    minPlayers: CREATE_TABLE_MIN_PLAYER_COUNT,
    maxPlayers: availablePlayers.length,
    maps: mapDefinitions,
    selection: buildSuggestedSelection_(normalizedMode, requestedCount)
  };
}

function createTable(selection, mode, playerCount, forceOverwrite) {
  if (typeof playerCount === "boolean" && forceOverwrite === undefined) {
    forceOverwrite = playerCount;
    playerCount = undefined;
  }

  clearAllCaches();
  const normalizedMode = normalizeCreateTableMode_(mode);
  const maps = getMaps().map(map => String(map)).filter(Boolean);
  const availablePlayers = getPlayers();
  const normalizedPlayerCount = normalizePlayerCount_(playerCount, availablePlayers.length);
  const normalizedSelection = normalizeTableSelection_(selection);

  validateTableSelection_(normalizedSelection, normalizedMode, maps);

  const activeRange = SpreadsheetApp.getActive().getActiveRange();
  if (!activeRange) throw new Error("Select a starting cell for the table.");

  const targetSheet = activeRange.getSheet();
  const startRow = activeRange.getRow();
  const startColumn = activeRange.getColumn();
  const height = normalizedSelection.length + 2;
  const width = normalizedPlayerCount + 2;

  if (startRow + height - 1 > targetSheet.getMaxRows()) {
    throw new Error("The target area exceeds the available rows.");
  }
  if (startColumn + width - 1 > targetSheet.getMaxColumns()) {
    throw new Error("The target area exceeds the available columns.");
  }

  const destination = targetSheet.getRange(startRow, startColumn, height, width);
  const hasContent = destination.getValues().some(row =>
    row.some(value => value !== "" && value !== null)
  );

  if (hasContent && !forceOverwrite) {
    throw new Error(
      "TARGET_NOT_EMPTY: the target area already contains data. Confirm to replace it."
    );
  }

  if (hasContent) {
    destination.breakApart();
    destination.clearContent();
  }

  const templateSheet = getSheet("Template");
  if (!templateSheet) throw new Error("The Template sheet was not found.");
  const tableHeader = getNextTableHeaderInfo_();

  copyTemplateRow_(
    templateSheet,
    targetSheet,
    TABLE_TEMPLATE_HEADER_ROW,
    startRow,
    startColumn,
    normalizedPlayerCount
  );
  targetSheet
    .getRange(startRow, startColumn)
    .setValue(`${tableHeader.number} - ${tableHeader.date}`);

  const mapValues = [];
  normalizedSelection.forEach((item, index) => {
    const targetRow = startRow + index + 1;
    const sourceRow = item.type === "NI"
      ? TABLE_TEMPLATE_NI_ROW
      : TABLE_TEMPLATE_SHROOM_ROW;

    copyTemplateRow_(
      templateSheet,
      targetSheet,
      sourceRow,
      targetRow,
      startColumn,
      normalizedPlayerCount
    );
    mapValues.push([
      getMapLabel_(item.map),
      ...Array(normalizedPlayerCount).fill(""),
      item.type
    ]);
  });

  const mapRange = targetSheet.getRange(
    startRow + 1,
    startColumn,
    normalizedSelection.length,
    width
  );
  mapRange.setValues(mapValues);
  mapRange.offset(0, 0, normalizedSelection.length, 1).clearDataValidations();

  removeInternalCategoryBorders_(targetSheet, startRow, startColumn, width, normalizedSelection);
  mergeCategoryCells_(targetSheet, startRow, startColumn, normalizedPlayerCount, normalizedSelection);

  const totalRow = startRow + normalizedSelection.length + 1;
  copyTemplateRow_(
    templateSheet,
    targetSheet,
    TABLE_TEMPLATE_TOTAL_ROW,
    totalRow,
    startColumn,
    normalizedPlayerCount
  );
  removeInternalPlayerBorders_(
    targetSheet,
    startRow,
    startColumn + 1,
    height,
    normalizedPlayerCount
  );
  targetSheet.getRange(totalRow, startColumn).setValue("Total");

  const firstMapRow = startRow + 1;
  const lastMapRow = totalRow - 1;
  const formulas = [];
  for (let offset = 1; offset <= normalizedPlayerCount; offset++) {
    const column = columnToLetter_(startColumn + offset);
    formulas.push(`=SUM_TIMES(${column}${firstMapRow}:${column}${lastMapRow})`);
  }

  targetSheet
    .getRange(totalRow, startColumn + 1, 1, normalizedPlayerCount)
    .setFormulas([formulas]);
  targetSheet.getRange(totalRow, startColumn + normalizedPlayerCount + 1).clearContent();

  return {
    sheetName: targetSheet.getName(),
    range: targetSheet.getRange(startRow, startColumn, height, width).getA1Notation(),
    startRow,
    startColumn,
    endRow: totalRow,
    count: normalizedSelection.length,
    playerCount: normalizedPlayerCount,
    mode: normalizedMode,
    selection: normalizedSelection
  };
}

function copyTemplateRow_(templateSheet, targetSheet, sourceRow, targetRow, targetColumn, playerCount) {
  const sourceMap = templateSheet.getRange(sourceRow, TABLE_TEMPLATE_START_COLUMN);
  const sourcePlayer = templateSheet.getRange(sourceRow, TABLE_TEMPLATE_START_COLUMN + 1);
  const sourceType = templateSheet.getRange(
    sourceRow,
    TABLE_TEMPLATE_START_COLUMN + TABLE_TEMPLATE_WIDTH - 1
  );
  const playerTarget = targetSheet.getRange(targetRow, targetColumn + 1, 1, playerCount);

  sourceMap.copyTo(targetSheet.getRange(targetRow, targetColumn), { contentsOnly: false });
  sourceType.copyTo(
    targetSheet.getRange(targetRow, targetColumn + playerCount + 1),
    { contentsOnly: false }
  );
  sourcePlayer.copyFormatToRange(
    targetSheet,
    targetColumn + 1,
    targetColumn + playerCount,
    targetRow,
    targetRow
  );

  if (sourceRow === TABLE_TEMPLATE_HEADER_ROW) {
    const validation = sourcePlayer.getDataValidation();
    if (validation) playerTarget.setDataValidation(validation);
    playerTarget.setValues([Array(playerCount).fill(sourcePlayer.getValue())]);
  }

  targetSheet.setRowHeight(targetRow, templateSheet.getRowHeight(sourceRow));
}

function removeInternalPlayerBorders_(sheet, startRow, startColumn, height, playerCount) {
  if (playerCount < 2) return;

  sheet
    .getRange(startRow, startColumn, height, playerCount)
    .setBorder(
      null,
      null,
      null,
      null,
      true,
      null,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );
}

function removeInternalCategoryBorders_(sheet, startRow, startColumn, width, selection) {
  let groupStart = 0;

  for (let index = 1; index <= selection.length; index++) {
    const isEndOfGroup = index === selection.length || selection[index].type !== selection[groupStart].type;
    if (!isEndOfGroup) continue;

    const groupLength = index - groupStart;
    if (groupLength > 1) {
      sheet
        .getRange(startRow + groupStart + 1, startColumn, groupLength, width)
        .setBorder(null, null, null, null, null, false);
    }

    groupStart = index;
  }
}

function mergeCategoryCells_(sheet, startRow, startColumn, playerCount, selection) {
  let groupStart = 0;

  for (let index = 1; index <= selection.length; index++) {
    const isEndOfGroup = index === selection.length || selection[index].type !== selection[groupStart].type;
    if (!isEndOfGroup) continue;

    const groupLength = index - groupStart;
    const categoryCell = sheet.getRange(
      startRow + groupStart + 1,
      startColumn + playerCount + 1,
      groupLength,
      1
    );

    if (groupLength > 1) categoryCell.merge();
    categoryCell
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    groupStart = index;
  }
}

function normalizeCreateTableMode_(mode) {
  if (mode === "NI" || mode === "Custom") return mode;
  return "Shroom";
}

function normalizeCreateTableCount_(count, mapCount) {
  const value = count === undefined || count === null || count === ""
    ? CREATE_TABLE_DEFAULT_COUNT
    : Number(count);

  if (!Number.isInteger(value)) throw new Error("The map count must be an integer.");
  if (value < 1 || value > Math.min(CREATE_TABLE_MAX_COUNT, mapCount)) {
    throw new Error(`The map count must be between 1 and ${Math.min(CREATE_TABLE_MAX_COUNT, mapCount)}.`);
  }

  return value;
}

function normalizePlayerCount_(count, availablePlayerCount) {
  if (availablePlayerCount < CREATE_TABLE_MIN_PLAYER_COUNT) {
    throw new Error("At least 2 players must be available in Template.");
  }

  const defaultCount = Math.min(CREATE_TABLE_DEFAULT_PLAYER_COUNT, availablePlayerCount);
  const value = count === undefined || count === null || count === ""
    ? defaultCount
    : Number(count);

  if (!Number.isInteger(value)) throw new Error("The player count must be an integer.");
  if (value < CREATE_TABLE_MIN_PLAYER_COUNT || value > availablePlayerCount) {
    throw new Error(
      `The player count must be between ${CREATE_TABLE_MIN_PLAYER_COUNT} and ${availablePlayerCount}.`
    );
  }

  return value;
}

function normalizeTableSelection_(selection) {
  if (!Array.isArray(selection) || selection.length === 0) {
    throw new Error("Select at least one map.");
  }

  return selection.map(item => {
    if (typeof item === "string") return { map: normalizeMapId_(item), type: "Shroom" };

    const map = item && item.map !== undefined ? normalizeMapId_(item.map) : "";
    const type = item && (item.type || item.format)
      ? String(item.type || item.format).trim()
      : "";

    return { map, type };
  });
}

function validateTableSelection_(selection, mode, maps) {
  const mapSet = new Set(maps);
  const seen = new Set();

  selection.forEach(item => {
    if (!mapSet.has(item.map)) throw new Error(`Unknown map: ${item.map || "(empty)"}.`);
    if (seen.has(item.map)) throw new Error(`Map ${item.map} is selected more than once.`);
    if (!CREATE_TABLE_FORMATS.includes(item.type)) {
      throw new Error(`Invalid format for ${item.map}.`);
    }
    seen.add(item.map);
  });

  if (mode === "Shroom" || mode === "NI") {
    if (selection.length !== CREATE_TABLE_DEFAULT_COUNT) {
      throw new Error(`${mode} requires exactly 5 maps.`);
    }

    const shroomCount = selection.filter(item => item.type === "Shroom").length;
    const expectedShroomCount = mode === "Shroom" ? 3 : 2;
    if (shroomCount !== expectedShroomCount) {
      throw new Error(
        `${mode} requires ${expectedShroomCount} Shroom maps and ${CREATE_TABLE_DEFAULT_COUNT - expectedShroomCount} NI maps.`
      );
    }
  }
}

/***********************
 * TABLE BLOCK DETECTION
 ***********************/

function getTableBlocks(dataValues, mapNames) {
  const maps = new Set((mapNames || getMaps()).map(map => normalizeMapId_(map)));
  const blocks = [];
  let current = null;

  for (let rowIndex = 0; rowIndex < dataValues.length; rowIndex++) {
    const row = dataValues[rowIndex] || [];

    if (current) {
      const label = getCellText_(row[current.startColumn]);
      const mapId = normalizeMapId_(label);

      if (label === "Total") {
        current.totalRow = rowIndex + 1;
        if (current.mapRows.length > 0) blocks.push(current);
        current = null;
        continue;
      }

      if (isMapCellValue_(label, maps)) {
        const formatInfo = findFormatInRow_(row);
        const previousFormat = current.mapRows.length > 0
          ? current.mapRows[current.mapRows.length - 1].format
          : "";
        const format = formatInfo ? formatInfo.format : previousFormat;
        current.mapRows.push({
          row: rowIndex + 1,
          column: current.startColumn,
          map: mapId,
          format,
          formatColumn: formatInfo ? formatInfo.column : null
        });
        if (formatInfo && !current.formatColumn) current.formatColumn = formatInfo.column;
        continue;
      }

      const replacementColumn = findTableHeaderColumn_(dataValues, rowIndex, maps);
      if (replacementColumn !== -1) {
        current = createTableBlock_(rowIndex + 1, replacementColumn);
      }
      continue;
    }

    const headerColumn = findTableHeaderColumn_(dataValues, rowIndex, maps);
    if (headerColumn !== -1) current = createTableBlock_(rowIndex + 1, headerColumn);
  }

  return blocks;
}

function createTableBlock_(headerRow, startColumn) {
  return {
    headerRow,
    totalRow: null,
    startColumn,
    formatColumn: null,
    mapRows: []
  };
}

function findTableHeaderColumn_(dataValues, rowIndex, maps) {
  const row = dataValues[rowIndex] || [];
  const lookAheadEnd = Math.min(dataValues.length, rowIndex + CREATE_TABLE_MAX_COUNT + 4);

  for (let column = 0; column < row.length; column++) {
    const value = getCellText_(row[column]);
    if (!value || value === "Total" || isMapCellValue_(value, maps)) continue;
    if (!row.slice(column + 1).some(isNonEmptyCell_)) continue;

    for (let nextRow = rowIndex + 1; nextRow < lookAheadEnd; nextRow++) {
      const nextValue = getCellText_((dataValues[nextRow] || [])[column]);
      if (isMapCellValue_(nextValue, maps)) return column;
      if (nextValue === "Total") break;
    }
  }

  return -1;
}

function isMapCellValue_(value, maps) {
  const text = getCellText_(value);
  const mapId = normalizeMapId_(text);
  if (!text || !maps.has(mapId)) return false;

  const suffix = text.replace(/^\d{1,2}\s*-\s*/, "").trim();
  if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(suffix)) return false;
  if (/^D{2}\/M{2}\/Y{4}$/i.test(suffix)) return false;
  return true;
}

function findFormatInRow_(row) {
  for (let column = 0; column < row.length; column++) {
    const value = getCellText_(row[column]);
    if (CREATE_TABLE_FORMATS.includes(value)) return { format: value, column };
  }
  return null;
}

function getActiveTableSelection() {
  clearAllCaches();
  const activeRange = SpreadsheetApp.getActive().getActiveRange();
  if (!activeRange) return { found: false, message: "No active cell." };

  const sheet = activeRange.getSheet();
  if (sheet.getName() !== "Tables") {
    return {
      found: false,
      message: "Select a cell inside a table on the Tables sheet."
    };
  }

  const data = sheet.getDataRange().getValues();
  const block = getTableBlocks(data).find(item =>
    activeRange.getRow() >= item.headerRow && activeRange.getRow() <= item.totalRow
  );

  if (!block) {
    return {
      found: false,
      message: "The active cell is not inside a recognized table."
    };
  }

  const headerCell = (data[block.headerRow - 1] || [])[block.startColumn];
  const tableHeader = parseTableHeaderInfo_(headerCell);

  return {
    found: true,
    sheetName: sheet.getName(),
    headerRow: block.headerRow,
    totalRow: block.totalRow,
    tableNumber: tableHeader.number,
    tableDate: tableHeader.date,
    startColumn: block.startColumn + 1,
    selection: block.mapRows.map(item => ({
      map: item.map,
      type: item.format || "Shroom"
    }))
  };
}

function getCurrentTableDate_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function parseTableHeaderInfo_(value) {
  const text = getCellText_(value);
  const match = text.match(/^\s*(\d+)\s*-\s*(\d{1,2}\/\d{1,2}\/\d{2,4})\s*$/);
  if (!match) return { number: "XX", date: getCurrentTableDate_(), recognized: false };

  return { number: match[1], date: match[2], recognized: true };
}

function getNextTableHeaderInfo_() {
  const tablesSheet = getSheet("Tables");
  if (!tablesSheet) return { number: "XX", date: getCurrentTableDate_() };

  const data = tablesSheet.getDataRange().getValues();
  const blocks = getTableBlocks(data);
  if (blocks.length === 0) return { number: "1", date: getCurrentTableDate_() };

  const lastBlock = blocks[blocks.length - 1];
  const headerCell = (data[lastBlock.headerRow - 1] || [])[lastBlock.startColumn];
  const lastHeader = parseTableHeaderInfo_(headerCell);
  if (!lastHeader.recognized) return { number: "XX", date: getCurrentTableDate_() };

  return {
    number: String(Number(lastHeader.number) + 1),
    date: getCurrentTableDate_()
  };
}

function getCellText_(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function isNonEmptyCell_(value) {
  return value !== "" && value !== null && value !== undefined;
}

/***********************
 * HISTORY / MAP DATA
 ***********************/

function getMapsHistory() {
  const data = getSheetData("Tables");
  return getTableBlocks(data).flatMap(block =>
    block.mapRows.map(item => ({ map: item.map, format: item.format }))
  );
}

function getLastFormatForMap(map, history) {
  for (let i = history.length - 1; i >= 0; i--) {
    if (history[i].map === map) return history[i].format || null;
  }
  return null;
}

function getWeightedMaps() {
  const definitions = getMapDefinitions();
  const defaultWeight = Math.max(1, ...definitions.map(map => map.weight || 0));
  return definitions
    .map(map => ({ ...map, weight: map.weight === null ? defaultWeight : map.weight }))
    .filter(map => map.weight > 0)
    .map(map => ({
      map: map.id,
      played: map.count,
      weight: map.weight
    }));
}

/***********************
 * CONTEXT CACHE
 ***********************/

function getContext() {
  if (_tablesCache) return _tablesCache;

  const history = getMapsHistory();
  const weightedMaps = getWeightedMaps();
  const excludedMaps = history.slice(-5).map(item => item.map);
  const candidates = weightedMaps.filter(item => !excludedMaps.includes(item.map));
  const totalWeight = candidates.reduce((sum, item) => sum + item.weight, 0);

  _tablesCache = { history, candidates, totalWeight };
  return _tablesCache;
}

function clearCache() {
  _tablesCache = null;
}

function pickFromPool(pool) {
  const index = Math.floor(Math.random() * pool.length);
  return pool.splice(index, 1)[0];
}

function pickWeightedFromPool_(pool) {
  const totalWeight = pool.reduce((sum, item) => sum + Math.max(0, item.weight), 0);
  if (totalWeight <= 0) return pickFromPool(pool);

  let target = Math.random() * totalWeight;
  for (let i = 0; i < pool.length; i++) {
    target -= Math.max(0, pool[i].weight);
    if (target < 0) return pool.splice(i, 1)[0];
  }

  return pool.pop();
}

/***********************
 * MAP GENERATION
 ***********************/

function generateMaps(count = CREATE_TABLE_DEFAULT_COUNT) {
  const requestedCount = normalizeCreateTableCount_(count, getMaps().length);
  const { candidates } = getContext();
  const selected = [];
  const selectedSet = new Set();

  drawMapsFromPool_(candidates, requestedCount, selected, selectedSet);

  // If the requested custom size includes recently played maps, complete the
  // draw from the remaining weighted maps before falling back to the dropdown.
  if (selected.length < requestedCount) {
    const fallbackWeighted = getWeightedMaps()
      .filter(item => !selectedSet.has(item.map));
    drawMapsFromPool_(fallbackWeighted, requestedCount, selected, selectedSet);
  }

  if (selected.length < requestedCount) {
    const fallbackMaps = getMaps()
      .map(map => ({ map: String(map), weight: 1 }))
      .filter(item => !selectedSet.has(item.map));
    drawMapsFromPool_(fallbackMaps, requestedCount, selected, selectedSet);
  }

  if (selected.length < requestedCount) {
    throw new Error(`Unable to draw ${requestedCount} unique maps.`);
  }

  return selected;
}

function drawMapsFromPool_(sourcePool, requestedCount, selected, selectedSet) {
  const pool = sourcePool
    .filter(item => item && item.map && !selectedSet.has(item.map))
    .map(item => ({ map: item.map, weight: Number(item.weight) || 0 }));

  while (selected.length < requestedCount && pool.length > 0) {
    const picked = pickWeightedFromPool_(pool);
    if (selectedSet.has(picked.map)) continue;
    selected.push(picked.map);
    selectedSet.add(picked.map);
  }
}

function buildSuggestedSelection_(mode, count) {
  const maps = generateMaps(count);
  if (mode === "Custom") return maps.map(map => ({ map, type: "Shroom" }));
  return assignFormats_(maps, mode);
}

function assignFormats_(maps, mode) {
  const { history } = getContext();
  let shroomSlots = mode === "Shroom" ? 3 : 2;
  let niSlots = mode === "Shroom" ? 2 : 3;

  return maps.map(map => {
    const lastFormat = getLastFormatForMap(map, history);
    const preferred = lastFormat === "Shroom" ? "NI" : "Shroom";
    let type;

    if (preferred === "Shroom" && shroomSlots > 0) {
      type = "Shroom";
      shroomSlots--;
    } else if (preferred === "NI" && niSlots > 0) {
      type = "NI";
      niSlots--;
    } else if (shroomSlots > 0) {
      type = "Shroom";
      shroomSlots--;
    } else {
      type = "NI";
      niSlots--;
    }

    return { map, type };
  }).sort((left, right) => {
    if (left.type === right.type) return 0;
    return left.type === "Shroom" ? -1 : 1;
  });
}

function testGenerateWithFormat(templateType) {
  clearCache();
  const mode = templateType === "NI" ? "NI" : "Shroom";
  return buildSuggestedSelection_(mode, CREATE_TABLE_DEFAULT_COUNT);
}

function testShroom() {
  return testGenerateWithFormat("Shroom");
}

function testNI() {
  return testGenerateWithFormat("NI");
}

/***********************
 * SHEET FUNCTION
 ***********************/

function openRefreshMapChances() {
  const html = HtmlService.createHtmlOutputFromFile("ui/RefreshMapsPopup")
    .setWidth(420)
    .setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, "Refresh Maps %");
}

function getRefreshMapOptions() {
  clearAllCaches();
  const tableBlocks = getTableBlocks(getSheetData("Tables"));
  return {
    defaultIgnoredTables: DEFAULT_IGNORED_TABLES_FOR_MAP_CHANCES,
    tableCount: tableBlocks.length
  };
}

function refreshMapChances(ignoredTables) {
  clearAllCaches();
  const mapDefinitions = getMapDefinitions();
  if (mapDefinitions.length === 0) throw new Error("No maps are available in Template.");

  const maps = mapDefinitions.map(map => map.id);
  const tableBlocks = getTableBlocks(getSheetData("Tables"), maps);
  const ignoredCount = normalizeIgnoredTableCount_(ignoredTables, tableBlocks.length);
  const countedTableBlocks = tableBlocks.slice(ignoredCount);
  const counts = new Map(maps.map(map => [map, 0]));

  countedTableBlocks.forEach(block => {
    block.mapRows.forEach(mapRow => {
      if (counts.has(mapRow.map)) counts.set(mapRow.map, counts.get(mapRow.map) + 1);
    });
  });

  const maxCount = Math.max(0, ...counts.values());
  const excluded = new Set(
    countedTableBlocks
      .flatMap(block => block.mapRows.map(mapRow => mapRow.map))
      .slice(-5)
  );
  const weights = mapDefinitions.map(map => maxCount - counts.get(map.id) + 1);
  const totalWeight = mapDefinitions.reduce((total, map, index) => {
    return total + (excluded.has(map.id) ? 0 : weights[index]);
  }, 0);
  const output = mapDefinitions.map((map, index) => [
    counts.get(map.id),
    weights[index],
    excluded.has(map.id) || totalWeight === 0 ? 0 : weights[index] / totalWeight * 100
  ]);

  const sheet = getSheet("Template");
  const outputRange = sheet.getRange(MAP_CATALOG_START_ROW, 4, output.length, 3);
  outputRange.setValues(output);
  outputRange.offset(0, 2, output.length, 1).setNumberFormat('0.00"%"');

  return {
    ignoredTables: ignoredCount,
    tableCount: tableBlocks.length,
    mapCount: mapDefinitions.length
  };
}

function normalizeIgnoredTableCount_(value, tableCount) {
  const normalized = value === undefined || value === null || value === ""
    ? DEFAULT_IGNORED_TABLES_FOR_MAP_CHANCES
    : Number(value);

  if (!Number.isInteger(normalized) || normalized < 0 || normalized > tableCount) {
    throw new Error(`The number of tables to ignore must be between 0 and ${tableCount}.`);
  }

  return normalized;
}

function columnToLetter_(column) {
  let result = "";
  let current = column;

  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }

  return result;
}

/***********************
 * COMPATIBILITY COPY
 ***********************/

function insertTemplate(inRange) {
  const ss = SpreadsheetApp.getActive();
  const targetRange = ss.getActiveRange();
  if (!targetRange) return;

  const templateSheet = ss.getSheetByName("Template");
  const templateRange = templateSheet.getRange(inRange);
  const destination = targetRange
    .getSheet()
    .getRange(
      targetRange.getRow(),
      targetRange.getColumn(),
      templateRange.getNumRows(),
      templateRange.getNumColumns()
    );

  templateRange.copyTo(destination, { contentsOnly: false });
}
