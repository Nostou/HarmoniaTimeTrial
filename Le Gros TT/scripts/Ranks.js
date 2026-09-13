function daysBetween(dateString) {
  const recordDate = new Date(dateString);
  const today = new Date();
  const diff = today - recordDate;
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

function stripHtml(str) {
  if (!str) return "";
  return str
    .replace(/<[^>]+>/g, "")
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
  return raw.split(/[\/／]/)[0].replace(/\s+/g, " ").trim();
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
    if (!tds || tds.length < 5) return;

    const map = stripHtml(tds[0]);
    const time = formatShroomTime(stripHtml(tds[1]));
    const player = stripHtml(tds[2]);
    const date = stripHtml(tds[4]);

    if (!map || !time || !player || !date) return;

    records.push({map, time, player, days: daysBetween(date)});
  });

  return records;
}

function fetchNIWR() {
  const ss = SpreadsheetApp.openById("1sK4dmfvS2db7grHRsUyFGcUx9UXVN6iV7a-kYlyz4xU");
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  const records = [];

  for (let row = 1; row < data.length; row++) {
    const rawMap = data[row][1]; // colonne B
    if (!rawMap || typeof rawMap !== "string") continue;

    const map = normalizeNIMapName(rawMap);
    
    const wrRow = row + 1;
    if (!data[wrRow]) continue;

    const timeRaw = data[wrRow][17]; // colonne K
    const player = data[wrRow][16];   // colonne J

    if (!timeRaw || !player) continue;

    const timeClean = normalizeWRTime(timeRaw.toString());
    records.push({map: normalizeNIMapName(rawMap), time: timeClean, player: player.toString().trim(), days: null});
  }

  return records;
}

/*function formatCell(record) {
  if (!record) return "";

  if (record.days === null) {
    return `${record.time}\n${record.player}`;
  }

  return `${record.time}\n${record.player} - ${record.days}d`;
}*/

function formatWR(records, expectedCount) {
  const values = [];

  for (let i = 0; i < expectedCount; i++) {
    const r = records[i];
    values.push([r.time]);
  }

  return values;
}

function logRecords(title, records, withDays) {
  Logger.log("=== %s ===", title);

  Object.keys(records)
    .forEach(map => {
      const r = records[map];
      if (withDays) {
        Logger.log("%s | %s | %s | %s days", map, r.time, r.player, r.days);
      } else {
        Logger.log("%s | %s | %s", map, r.time, r.player);
      }
    });

  Logger.log("=== END %s ===", title);
}

function fillRankTargets() {
  const sheet = getSheet("Ranks");
  const maps = getMaps();

  const multipliers = getRankMultipliers();

  // Lecture WR + noms des ranks
  const wrValues = sheet.getRange(2, 3, maps.length * 2 + 1, 1).getValues(); // colonne C
  const rankHeaders = sheet.getRange(1, 4, 1, 15).getValues()[0]; // D → R

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
    .getRange(2, 4, output.length, output[0].length)
    .setValues(output);
}

function refreshRanks() {
  //if (!assertOwnerOnly()) return;

  initMaps("Ranks");

  const ranks = getSheet("Ranks");
  const maps = getMaps();

  const shroomWR = fetchShroomWR();
  const niWR = fetchNIWR();

  //logRecords("SHROOM WR", shroomWR, true);
  //logRecords("NI WR", niWR, false);

  const shroomStart = 2;
  const niStart = maps.length + 3;

  ranks
    .getRange(shroomStart, 3, maps.length, 1)
    .setValues(formatWR(shroomWR, maps.length));

  ranks
    .getRange(niStart, 3, maps.length, 1)
    .setValues(formatWR(niWR, maps.length));

  fillRankTargets();
}
