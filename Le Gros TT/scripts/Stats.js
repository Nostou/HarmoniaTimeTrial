// ================================================================
// STATS HR - Calcul des statistiques des joueurs
// ================================================================

// Couleur correspondant à une victoire dans Tables HR
const COULEUR_VICTOIRE = "#ffe599";

// Blocs dont la distance entre l'en-tête et le Total n'est pas 6
const EXCEPTIONS_DISTANCE = {
  77: 4,
  332: 5
};

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

function getTotalRows(dataValues) {
  const totalRows = [];
  for (let r = 0; r < dataValues.length; r++) {
    if (dataValues[r][0] === "Total") totalRows.push(r + 1);
  }
  return totalRows;
}

// ----------------------------------------------------------------
// APPARITIONS, VICTOIRES
// ----------------------------------------------------------------
function getApparitionsStats(players, dataValues, dataColors) {
  const stats = {};
  players.forEach(p => {
    stats[p] = { apparitions: 0, victoires: 0 };
  });

  const totalRows = getTotalRows(dataValues);

  totalRows.forEach(totalRow => {
    const distance = EXCEPTIONS_DISTANCE[totalRow] || 6;
    const headerRow = totalRow - distance;
    const headerValues = dataValues[headerRow - 1];

    for (let c = 1; c < headerValues.length; c++) {
      const player = headerValues[c];
      if (!player || !stats[player]) continue;

      stats[player].apparitions++;

      const color = dataColors[totalRow - 1][c].toLowerCase();
      if (color === COULEUR_VICTOIRE) {
        stats[player].victoires++;
      }
    }
  });

  return stats;
}

// ----------------------------------------------------------------
// MAPS JOUÉES (séparées Shroom / Nita)
// ----------------------------------------------------------------
function getNBPlayedTracks(players, headers, dataShrooms, dataNita) {
  const mapsJouees = {};

  players.forEach(player => {
    const colIndex = headers.indexOf(player);

    if (colIndex === -1) {
      mapsJouees[player] = "0 - 0";
      return;
    }

    function compterTirets(data) {
      let nbTirets = 0;
      data.forEach(row => {
        const valeur = row[2 + colIndex].toString().split("\n")[0].trim();
        if (valeur === "-") nbTirets++;
      });
      return nbTirets;
    }

    const shroomJouees = dataShrooms.length - compterTirets(dataShrooms);
    const nitaJouees    = dataNita.length - compterTirets(dataNita);

    mapsJouees[player] = `${shroomJouees} - ${nitaJouees}`;
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
    bestTimes[p] = `${bestTimesShroom[p]} - ${bestTimesNita[p]}`;
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tablesSheet   = ss.getSheetByName("Tables");
  const statsSheet    = ss.getSheetByName("Stats");
  const archivesSheet = ss.getSheetByName("Archives");
  const ranksSheet    = ss.getSheetByName("Ranks");

  const players = statsSheet.getRange("A2:A").getValues()
    .map(r => r[0]).filter(v => v !== "");

  const lastRow    = tablesSheet.getLastRow();
  const lastCol    = tablesSheet.getLastColumn();
  const dataValues = tablesSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const dataColors = tablesSheet.getRange(1, 1, lastRow, lastCol).getBackgroundColors();

  const headers     = archivesSheet.getRange("D1:M1").getDisplayValues()[0];
  const dataShrooms = archivesSheet.getRange("B2:M31").getDisplayValues();
  const dataNita    = archivesSheet.getRange("B33:M62").getDisplayValues();

  const wrShrooms = ranksSheet.getRange("C2:C31").getDisplayValues().map(r => r[0]);
  const wrNita    = ranksSheet.getRange("C33:C62").getDisplayValues().map(r => r[0]);

  const apparitionsStats = getApparitionsStats(players, dataValues, dataColors);
  const bestTimesStats   = getBestTimesStats(players, headers, dataShrooms, dataNita);
  const nbTracksPlayed   = getNBPlayedTracks(players, headers, dataShrooms, dataNita);
  const rangMoyenStats   = getRangMoyenStats(players, headers, dataShrooms, dataNita, wrShrooms, wrNita);

  const pct = (val, total) => total > 0 ? Math.round(val / total * 100) : 0;

  const output = players.map(p => {
    const a = apparitionsStats[p];
    return [
      a.apparitions,
      `${a.victoires} (${pct(a.victoires, a.apparitions)}%)`,
      nbTracksPlayed[p],
      bestTimesStats[p],
      rangMoyenStats[p]
    ];
  });

  statsSheet.getRange(2, 2, output.length, 5).setValues(output);
  Logger.log("Stats calculées pour " + players.length + " joueurs.");
}