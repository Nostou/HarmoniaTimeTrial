function stripHtml(str) {
  if (!str) return "";
  return str
    .replace(/<[^>]+>/g, "")
    .replace(/&(apos|rsquo|#0*39|#x27);/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/&amp;/gi, "&")
    .replace(/&nbsp;/g, " ")
    .trim();
}

function formatShroomTime(raw) {
  if (!raw) return "";

  const match = raw.match(/(\d+)'(\d+)"(\d+)/);
  if (!match) return raw;

  const minutes = match[1];
  const seconds = match[2].padStart(2, "0");
  const millis = match[3].padStart(3, "0");

  return `${minutes}:${seconds}.${millis}`;
}

function normalizeNIMapName(raw) {
  if (!raw || typeof raw !== "string") return "";
  return raw.split(/[\/／]/)[0].replace(/\s*\([^()]*\)\s*$/, "").replace(/\s+/g, " ").trim();
}

function normalizeWRTime(t) {
  if (typeof t !== "string") return null;

  t = t.trim().replace('*', '');

  // m:ss.fff
  if (/^\d+:[0-5]\d\.\d{3}$/.test(t)) return t;

  // faute connue m:ss:fff
  if (/^\d+:[0-5]\d:\d{3}$/.test(t)) {
    return t.replace(/:(\d{3})$/, ".$1");
  }

  return null;
}

function fetchShroomWR() {
  const url = "https://mkwrs.com/mkworld/wrs.php";
  const html = UrlFetchApp.fetch(url).getContentText();

  const records = [];

  const rows = html.match(/<tr[\s\S]*?<\/tr>/g);
  if (!rows) return records;

  rows.forEach(row => {
    if (!row.includes("display.php?track=")) return;

    const tds = row.match(/<td[\s\S]*?<\/td>/g);
    if (!tds || tds.length < 3) return;

    const map = stripHtml(tds[0]);
    const time = formatShroomTime(stripHtml(tds[1]));
    const player = stripHtml(tds[2]);

    if (!map || !time || !player) return;

    records.push({map, time, player});
  });

  return records;
}

function fetchNIWR() {
  const url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRy9H_mG4dlna5tMNoNjAnP9Fk26D0jAA2KFj6ZV37uKvi79sNqUBUrx29jEykMfsoBVLX1WJmLI_SS/pub?gid=0&single=true&output=csv";
  const data = Utilities.parseCsv(UrlFetchApp.fetch(url).getContentText("UTF-8"));
  const records = [];

  for (let row = 0; row < data.length; row++) {
    const metaColumn = data[row].findIndex(cell => String(cell).trim() === "Meta Only");
    if (metaColumn < 0) continue;
    const map = normalizeNIMapName(data[row][1]);
    const top = data[row + 1] || [];
    if (!map || String(top[metaColumn]).trim() !== "1") continue;
    const time = normalizeWRTime(String(top[metaColumn + 5] || ""));
    if (time) records.push({ map, time, player: String(top[metaColumn + 4] || "").trim() });
  }

  if (records.length === 0) throw new Error("No NI Meta Only records found in the published leaderboard.");
  return records;
}

function normalizeRankMapName(raw) {
  return String(raw || "")
    .toLowerCase()
    .replace(/[’']/g, "")
    .replace(/^snes\s+/, "")
    .replace(/[^a-z0-9]+/g, "");
}

function formatWR(records, mapDefinitions) {
  const timesByMap = new Map();

  records.forEach(record => {
    const key = normalizeRankMapName(record.map);
    if (key && !timesByMap.has(key)) timesByMap.set(key, record.time || "-");
  });

  return mapDefinitions.map(map => [
    timesByMap.get(normalizeRankMapName(map.name)) || "-"
  ]);
}

function fillRankTargets() {
  const sheet = getSheet("Ranks");
  const maps = getMaps();

  const multipliers = getRankMultipliers();

  // Lecture WR + noms des ranks
  const wrValues = sheet.getRange(2, 4, maps.length, 1).getValues()
    .concat(sheet.getRange(maps.length + 3, 4, maps.length, 1).getValues());
  const rankHeaders = sheet.getRange(1, 5, 1, 15).getValues()[0];

  // Préparer tableau résultat
  const output = [];

  for (let i = 0; i < wrValues.length; i++) {
    const wr = wrValues[i][0];
    const row = [];

    const wrMs = TIME_TO_MS(wr);
    if (wrMs === null) {
      row.push(...Array(rankHeaders.length).fill("-"));
      output.push(row);
      continue;
    }

    rankHeaders.forEach(rank => {
      const factor = multipliers[rank];
      if (!factor) {
        row.push("-");
      } else {
        row.push(MS_TO_TIME(wrMs * factor));
      }
    });

    output.push(row);
  }

  // Écriture en une fois
  sheet
    .getRange(2, 5, maps.length, output[0].length)
    .setValues(output.slice(0, maps.length));
  sheet.getRange(maps.length + 3, 5, maps.length, output[0].length)
    .setValues(output.slice(maps.length));

  const rankPalette = sheet.getRange(2, 5, 1, rankHeaders.length).getBackgrounds()[0];
  const mapBackgrounds = Array.from({ length: maps.length }, () => rankPalette.slice());
  sheet
    .getRange(2, 5, maps.length, rankHeaders.length)
    .setBackgrounds(mapBackgrounds);
  sheet
    .getRange(maps.length + 3, 5, maps.length, rankHeaders.length)
    .setBackgrounds(mapBackgrounds);
}

function refreshRanks() {
  //if (!assertOwnerOnly()) return;

  clearAllCaches();
  const ranks = getSheet("Ranks");
  const mapDefinitions = getMapDefinitions();
  const maps = mapDefinitions.map(map => map.id);

  const shroomWR = fetchShroomWR();
  const niWR = fetchNIWR();

  if (ranks.getRange(maps.length + 3, 1).getDisplayValue() !== "NI") {
    throw new Error("The NI section in Ranks does not match the map count. Adjust the section and Base/DLC labels before refreshing.");
  }
  initMaps("Ranks");

  const shroomStart = 2;
  const niStart = maps.length + 3;

  ranks
    .getRange(shroomStart, 4, maps.length, 1)
    .setValues(formatWR(shroomWR, mapDefinitions));

  ranks
    .getRange(niStart, 4, maps.length, 1)
    .setValues(formatWR(niWR, mapDefinitions));

  fillRankTargets();
}
