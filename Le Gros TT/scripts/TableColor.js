function colorTable(range, mapName, isShroom, rankCache) {
  const values = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();
  const richValues = range.getRichTextValues();

  const colors = ["#FFE599", "#D9D9D9", "#F9CB9C"]; // top 1 / 2 / 3

  for (let r = 0; r < values.length; r++) {
    const parsed = values[r]
      .map((v, c) => {
        const time = typeof v === "string" ? v.split("\n")[0].trim() : v;
        return { col: c, ms: TIME_TO_MS(time), text: time };
      })
      .filter(o => o.ms !== null);

    parsed.sort((a, b) => a.ms - b.ms);

    let rankIndex = 0;
    let last = null;

    parsed.forEach(o => {
      if (last === null || o.ms !== last) rankIndex++;
      last = o.ms;

      // ---------- Couleur (TOP 3 SEULEMENT) ----------
      if (rankIndex <= 3) {
        backgrounds[r][o.col] = colors[rankIndex - 1];
      } else {
        backgrounds[r][o.col] = "#FFFFFF";
      }

      if (mapName === "Total") {
        richValues[r][o.col] = makePlainRichText(o.text);
        return;
      }

      // ---------- Rank depuis la table Ranks ----------
      const rankInfo = getRank(rankCache, mapName, o.text, isShroom);
      const rankName = rankInfo ? rankInfo.rank : "Unranked";

      richValues[r][o.col] = buildTimeRankRichText(o.text, rankName);
    });
  }

  range.setRichTextValues(richValues);
  range.setBackgrounds(backgrounds);
}

function colorSelectedTable() {
  const range = SpreadsheetApp.getActive().getActiveRange();
  const sheet = range.getSheet();

  const values = range.getDisplayValues();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const rankCache = buildRankCache();

  let currentType = null;

  for (let r = 0; r < numRows; r++) {
    const row = values[r];

    const mapName = row[0];
    const typeCell = row[numCols - 1];

    // Mise à jour du type si présent
    if (typeCell === "Shroom" || typeCell === "NI") {
      currentType = typeCell;
    }

    if (!mapName || !currentType) continue;

    const isShroom = currentType === "Shroom";

    const timeRange = sheet.getRange(
      startRow + r,
      startCol + 1,
      1,
      numCols - 2
    );

    colorTable(timeRange, mapName, isShroom, rankCache);
  }
}
