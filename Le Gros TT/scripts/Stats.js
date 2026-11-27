// ================================================================
// STATS HR - Calcul des statistiques des joueurs
// ================================================================

// Couleur correspondant à une victoire dans Tables HR
const COULEUR_VICTOIRE = "#ffe599";

// Table de correspondance rapport temps/WR → rang
const RANK_TABLE = [
  { rank: "Legend",   value: 1.005 },
  { rank: "Master",   value: 1.010 },
  { rank: "Diamond",  value: 1.020 },
  { rank: "Ruby",     value: 1.030 },
  { rank: "Sapphire", value: 1.040 },
  { rank: "Emerald",  value: 1.050 },
  { rank: "Platinum", value: 1.060 },
  { rank: "Gold 1",   value: 1.070 },
  { rank: "Gold 2",   value: 1.080 },
  { rank: "Silver 1", value: 1.090 },
  { rank: "Silver 2", value: 1.100 },
  { rank: "Bronze 1", value: 1.110 },
  { rank: "Bronze 2", value: 1.120 },
  { rank: "Iron 1",   value: 1.135 },
  { rank: "Iron 2",   value: 1.500 },
];

// ----------------------------------------------------------------
// UTILITAIRES
// ----------------------------------------------------------------

function rapportToRank(rapport) {
  for (const r of RANK_TABLE) {
    if (rapport <= r.value) return r.rank;
  }
  return "Unranked";
}

function removeLegacyRankColors_(sheet) {
  const legacy = new Set(["Iron", "Bronze", "Silver", "Gold", "Platinum", "Emerald"]);
  const rules = sheet.getConditionalFormatRules();
  const remaining = rules.filter(rule => {
    const ranges = rule.getRanges();
    const condition = rule.getBooleanCondition();
    if (ranges.length !== 1 || !condition) return true;
    const range = ranges[0];
    return !(range.getA1Notation() === "F2:F10" &&
      condition.getCriteriaType() === SpreadsheetApp.BooleanCriteria.TEXT_CONTAINS &&
      legacy.has(String(condition.getCriteriaValues()[0])));
  });
  if (remaining.length !== rules.length) sheet.setConditionalFormatRules(remaining);
}

// ----------------------------------------------------------------
// APPARITIONS, VICTOIRES
// ----------------------------------------------------------------
function getApparitionsStats(players, dataValues, dataColors) {
  const stats = {};
  players.forEach(p => {
    stats[p] = { apparitions: 0, victoires: 0 };
  });

  const tableBlocks = getTableBlocks(dataValues);

  tableBlocks.forEach(block => {
    const headerValues = dataValues[block.headerRow - 1] || [];
    const totalColors = dataColors[block.totalRow - 1] || [];
    const playerEnd = block.formatColumn || headerValues.length;

    for (let c = block.startColumn + 1; c < playerEnd; c++) {
      const player = headerValues[c];
      if (!player || !stats[player]) continue;

      stats[player].apparitions++;

      const color = String(totalColors[c] || "").toLowerCase();
      if (color === COULEUR_VICTOIRE) {
        stats[player].victoires++;
      }
    }
  });

  return stats;
}

// ----------------------------------------------------------------
// MAPS JOUÉES (Shroom et NI comptent séparément)
// ----------------------------------------------------------------
function getNBPlayedTracks(players, headers, dataShrooms, dataNita) {
  const mapsJouees = {};

  players.forEach(player => {
    const colIndex = headers.indexOf(player);

    if (colIndex === -1) {
      mapsJouees[player] = `0/${dataShrooms.length + dataNita.length}`;
      return;
    }

    const played = [...dataShrooms, ...dataNita]
      .filter(row => TIME_TO_MS(row[2 + colIndex]) !== null).length;
    mapsJouees[player] = `${played}/${dataShrooms.length + dataNita.length}`;
  });

  return mapsJouees;
}


// ----------------------------------------------------------------
// MEILLEURS TEMPS (séparés Shroom / Nita)
// ----------------------------------------------------------------
function getBestTimesStats(players, headers, dataShrooms, dataNita) {
  const bestTimesShroom = {};
  const bestTimesNita   = {};

  players.forEach(p => {
    bestTimesShroom[p] = 0;
    bestTimesNita[p]   = 0;
  });

  function traiterBloc(data, target) {
    data.forEach(row => {
      const temps = row.slice(2);
      let bestMs = Infinity;
      let bestJoueur = "";

      temps.forEach((t, i) => {
        const clean = t.split("\n")[0].trim();
        if (!clean || clean === "-") return;
        const ms = TIME_TO_MS(clean);
        if (ms !== null && ms < bestMs) {
          bestMs = ms;
          bestJoueur = headers[i];
        }
      });

      if (bestJoueur && target[bestJoueur] !== undefined) {
        target[bestJoueur]++;
      }
    });
  }

  traiterBloc(dataShrooms, bestTimesShroom);
  traiterBloc(dataNita, bestTimesNita);

  const bestTimes = {};
  players.forEach(p => {
    bestTimes[p] = bestTimesShroom[p] + bestTimesNita[p];
  });

  return bestTimes;
}

// ----------------------------------------------------------------
// RANG MOYEN
// ----------------------------------------------------------------
function getRangMoyenStats(players, headers, dataShrooms, dataNita, wrShrooms, wrNita) {
  const rangs = {};

  players.forEach(player => {
    const colIndex = headers.indexOf(player);
    if (colIndex === -1) {
      rangs[player] = "N/A";
      return;
    }

    let totalRapport = 0;
    let count = 0;

    function traiterBloc(data, wrs) {
      data.forEach((row, i) => {
        const msTemps = TIME_TO_MS(row[2 + colIndex]);
        const msWR    = TIME_TO_MS(wrs[i]);
        if (msTemps === null || msWR === null || msWR === 0) return;
        totalRapport += msTemps / msWR;
        count++;
      });
    }

    traiterBloc(dataShrooms, wrShrooms);
    traiterBloc(dataNita, wrNita);

    if (count === 0) {
      rangs[player] = "N/A";
      return;
    }

    const moyenneRapport = totalRapport / count;
    Logger.log(`${player} : ${moyenneRapport.toFixed(4)} → ${rapportToRank(moyenneRapport)}`);
    rangs[player] = rapportToRank(moyenneRapport);
  });

  return rangs;
}

// ----------------------------------------------------------------
// FONCTION PRINCIPALE
// ----------------------------------------------------------------
function Stats() {
  clearAllCaches();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tablesSheet   = ss.getSheetByName("Tables");
  const statsSheet    = ss.getSheetByName("Stats");
  const archivesSheet = ss.getSheetByName("Archives");
  const ranksSheet    = ss.getSheetByName("Ranks");
  const maps           = getMaps();
  const players = getPlayers();
  const statsLastRow = statsSheet.getLastRow();

  const lastRow    = tablesSheet.getLastRow();
  const lastCol    = tablesSheet.getLastColumn();
  const dataValues = tablesSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const dataColors = tablesSheet.getRange(1, 1, lastRow, lastCol).getBackgroundColors();

  const headers = players.length
    ? archivesSheet.getRange(1, 5, 1, players.length).getDisplayValues()[0]
    : [];
  if (players.some((player, index) => headers[index] !== player)) {
    throw new Error("The player list in Archives is out of date. Run Custom > Refresh Archives first.");
  }
  const archiveWidth = players.length + 2;
  const dataShrooms = archivesSheet
    .getRange(2, 3, maps.length, archiveWidth)
    .getDisplayValues();
  const dataNita = archivesSheet
    .getRange(maps.length + 3, 3, maps.length, archiveWidth)
    .getDisplayValues();

  if (maps.some((map, index) => normalizeMapId_(dataShrooms[index][0]) !== map ||
      normalizeMapId_(dataNita[index][0]) !== map)) {
    throw new Error("The map list in Archives is out of date. Run Custom > Refresh Archives first.");
  }
  buildRankCache();

  const wrShrooms = ranksSheet
    .getRange(2, 4, maps.length, 1)
    .getDisplayValues()
    .map(row => row[0]);
  const wrNita = ranksSheet
    .getRange(maps.length + 3, 4, maps.length, 1)
    .getDisplayValues()
    .map(row => row[0]);

  const apparitionsStats = getApparitionsStats(players, dataValues, dataColors);
  const bestTimesStats   = getBestTimesStats(players, headers, dataShrooms, dataNita);
  const nbTracksPlayed   = getNBPlayedTracks(players, headers, dataShrooms, dataNita);
  const rangMoyenStats   = getRangMoyenStats(players, headers, dataShrooms, dataNita, wrShrooms, wrNita);
  const paletteRange = ranksSheet.getRange("U3:U17");
  const rankNames = paletteRange.getDisplayValues();
  const rankColors = paletteRange.getBackgrounds();
  const colorsByRank = new Map(rankNames.map((row, i) => [row[0], rankColors[i][0]]));

  const pct = (val, total) => total > 0 ? Math.round(val / total * 100) : 0;

  const output = players.map(p => {
    const a = apparitionsStats[p];
    return [
      p,
      a.apparitions,
      `${a.victoires} (${pct(a.victoires, a.apparitions)}%)`,
      nbTracksPlayed[p],
      bestTimesStats[p],
      rangMoyenStats[p]
    ];
  });

  const neededRows = players.length + 1;
  if (neededRows > statsSheet.getMaxRows()) statsSheet.insertRowsAfter(statsSheet.getMaxRows(), neededRows - statsSheet.getMaxRows());
  if (statsLastRow > 1) statsSheet.getRange(2, 1, statsLastRow - 1, 6).clearContent();
  removeLegacyRankColors_(statsSheet);
  if (statsLastRow > 1) statsSheet.getRange(2, 6, statsLastRow - 1, 1).setBackground("#ffffff");
  statsSheet.getRange(1, 4, 1, 2).setValues([["Track Progress", "Records"]]);
  if (output.length > 0) {
    statsSheet.getRange(2, 4, output.length, 1).setNumberFormat("@");
    statsSheet.getRange(2, 1, output.length, 6).setValues(output);
    statsSheet.getRange(2, 6, output.length, 1)
      .setBackgrounds(players.map(player => [colorsByRank.get(rangMoyenStats[player]) || "#ffffff"]));
  }
  Logger.log("Stats calculated for " + players.length + " players.");
}
