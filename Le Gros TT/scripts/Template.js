/***********************
 * TEMPLATE INSERTION
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

function insertTemplateShroom() {
  clearUILogs();
  insertTemplate("A1:G7");
  testShroom();
  showUILogs();
}

function insertTemplateNI() {
  clearUILogs();
  insertTemplate("A9:G15");
  testNI();
  showUILogs();
}

/***********************
 * DATA ACCESS
 ***********************/
function getMapsHistory() {
  const data = getSheetData("Tables");
  const maps = getMaps();

  const history = [];
  let format = "";

  for (let i = 0; i < data.length; i++) {

    const map = data[i][0];
    if (!maps.includes(map)) continue;

    for (let j = 2; j < data[i].length; j++) {
      const cell = data[i][j];
      if (cell === "") break;
      if (cell === "Shroom" || cell === "NI") format = cell;
    }

    history.push({ map, format });
  }

  return history;
}

function getLastPlayedMaps(count) {
  const data = getSheetData("Tables");
  const maps = getMaps();

  const result = [];
  for (let i = data.length-1; i >= 0 && result.length < count; i--) {
    let potentialMap = data[i][0];
    if (maps.includes(potentialMap)) {
      result.push(potentialMap);
    }
  }

  return result;
}

function getLastFormatForMap(map, history) {
  for (let i = history.length - 1; i >= 0; i--) {
    if (history[i].map === map) return history[i].format;
  }
  return null;
}

function getWeightedMaps() {
  const sheet = getSheet("Template");
  const data = sheet.getRange("I2:L31").getValues();

  return data
    .filter(r => r[0] && r[2] > 0)
    .map(r => ({
      map: r[0],
      played: r[1],
      weight: r[2]
    }));
}

/***********************
 * CONTEXT CACHE (CORE)
 ***********************/
let _tablesCache = null;

function getContext() {
  if (_tablesCache) return _tablesCache;

  const history = getMapsHistory();
  const weightedMaps = getWeightedMaps();

  const excludedMaps = history.slice(-5).map(h => h.map);

  const candidates = weightedMaps.filter(
    m => m.weight > 0 && !excludedMaps.includes(m.map)
  );

  const totalWeight = candidates.reduce((s, c) => s + c.weight, 0);

  _tablesCache = { history, candidates, totalWeight };
  return _tablesCache;
}

function clearCache() {
  _tablesCache = null;
}

/***********************
 * RANDOM / PICK LOGIC
 ***********************/
function pickFromPool(pool) {
  const index = Math.floor(Math.random() * pool.length);
  const picked = pool[index];

  for (let i = pool.length - 1; i >= 0; i--) {
    if (pool[i] === picked) pool.splice(i, 1);
  }

  return picked;
}

/***********************
 * MAP GENERATION
 ***********************/
function generateMaps() {
  const { candidates } = getContext();

  const pool = [];
  candidates.forEach(c => {
    for (let i = 0; i < c.weight; i++) {
      pool.push(c.map);
    }
  });

  const selected = [];
  while (selected.length < 5 && pool.length > 0) {
    selected.push(pickFromPool(pool));
  }

  return selected;
}

/***********************
 * SHEET FUNCTION
 ***********************/
function refreshMapChances() {
  //if (!assertOwnerOnly()) return;
  
  const sheet = getSheet("Template");
  const data = sheet.getRange("I2:K31").getValues();
  const excluded = new Set(getLastPlayedMaps(5));

  let totalWeight = 0;
  const weights = [];

  data.forEach(([map, , weight]) => {
    if (!map || weight <= 0 || excluded.has(map)) {
      weights.push(0);
    } else {
      weights.push(weight);
      totalWeight += weight;
    }
  });

  const chances = weights.map(w =>
    totalWeight ? [w / totalWeight] : [0]
  );

  sheet.getRange("L2:L31").setValues(chances);
}

/***********************
 * FORMAT ASSIGNMENT
 ***********************/

function testGenerateWithFormat(templateType) {
  clearCache();

  logUI(`GENERATE (${templateType})`);

  const maps = generateMaps();
  const { history, candidates, totalWeight } = getContext();

  let shroomSlots = templateType === "Shroom" ? 3 : 2;
  let niSlots     = templateType === "Shroom" ? 2 : 3;

  const result = [];
  const shroomLogs = [];
  const niLogs = [];

  maps.forEach((map, index) => {
    const lastFormat = getLastFormatForMap(map, history);
    const preferred = lastFormat === "Shroom" ? "NI" : "Shroom";

    let chosen;

    if (preferred === "Shroom" && shroomSlots > 0) {
      chosen = "Shroom";
      shroomSlots--;
    }
    else if (preferred === "NI" && niSlots > 0) {
      chosen = "NI";
      niSlots--;
    }
    else if (shroomSlots > 0) {
      chosen = "Shroom";
      shroomSlots--;
    }
    else {
      chosen = "NI";
      niSlots--;
    }

    const candidate = candidates.find(c => c.map === map);
    const chance = totalWeight > 0 ? (candidate.weight / totalWeight) * 100 : 0;

    //logUI(`${index + 1}. ${map} | last=${lastFormat || "-"} → ${chosen} | ${chance.toFixed(1)}%`);

    const line = `${map} | ${chosen} | ${chance.toFixed(1)}%`;
    if (chosen === "Shroom") shroomLogs.push(line);  
    else niLogs.push(line);

    result.push({ map, format: chosen });
  });

  let displayIndex = 1;

  function logGroup(group) {
    group.forEach(line => {
      logUI(`${displayIndex}. ${line}`);
      displayIndex++;
    });
  }

  logGroup(shroomLogs);
  logGroup(niLogs);

  return result;
}

function testShroom() {
  testGenerateWithFormat("Shroom");
}

function testNI() {
  testGenerateWithFormat("NI");
}