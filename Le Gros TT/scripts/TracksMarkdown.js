// Kept for compatibility with the former custom-menu entry.
function TracksMarkdown() {
  CreateTable();
}

function CreateTracksMarkdown(selection, selectedNita, tableInfo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Archives");
  const maps = getMaps();
  const players = getPlayers();
  const normalizedSelection = normalizeMarkdownSelection_(selection, selectedNita, maps);
  const header = getMarkdownHeaderInfo_(tableInfo);

  if (normalizedSelection.length === 0) {
    throw new Error("Select at least one track to generate Markdown.");
  }

  const headers = players.length
    ? sheet.getRange(1, 5, 1, players.length).getDisplayValues()[0]
    : [];
  const dataWidth = players.length + 2;
  const dataShrooms = sheet.getRange(2, 3, maps.length, dataWidth).getDisplayValues();
  const dataNita = sheet
    .getRange(maps.length + 3, 3, maps.length, dataWidth)
    .getDisplayValues();

  const shrooms = normalizedSelection.filter(item => !item.nita);
  const nita = normalizedSelection.filter(item => item.nita);

  function formatLigne(row) {
    const cours = row[0] || "";
    const moyenne = String(row[1] || "-").split("\n")[0].trim();
    const temps = row.slice(2);

    let bestMs = Infinity;
    let bestTime = "";
    let bestJoueur = "";

    temps.forEach((time, index) => {
      const clean = String(time || "").split("\n")[0].trim();
      if (!clean || clean === "-") return;

      const milliseconds = TIME_TO_MS(clean);
      if (milliseconds !== null && milliseconds < bestMs) {
        bestMs = milliseconds;
        bestTime = clean;
        bestJoueur = headers[index] || "";
      }
    });

    return `> [${cours}]\nAverage : ${moyenne}\nBest : ${bestTime} (${bestJoueur})\n`;
  }

  let text = `\`\`\`Gros TT n°${header.number} - ${header.date}\`\`\`\n`;
  text += "## Shrooms\n```\n";

  shrooms.forEach(item => {
    text += formatLigne(dataShrooms[item.index]) + "\n";
  });

  text += "```";

  if (nita.length > 0) {
    text += "\n## NITA\n```\n";
    nita.forEach(item => {
      text += formatLigne(dataNita[item.index]) + "\n";
    });
    text += "```";
  }

  return text;
}

function getMarkdownHeaderInfo_(tableInfo) {
  if (tableInfo && tableInfo.number && tableInfo.date) {
    return {
      number: String(tableInfo.number),
      date: String(tableInfo.date)
    };
  }

  const activeTable = getActiveTableSelection();
  if (activeTable.found) {
    return {
      number: activeTable.tableNumber,
      date: activeTable.tableDate
    };
  }

  return getNextTableHeaderInfo_();
}

function normalizeMarkdownSelection_(selection, selectedNita, maps) {
  if (!Array.isArray(selection)) return [];

  const legacyNita = new Set(Array.isArray(selectedNita) ? selectedNita : []);
  const seen = new Set();

  return selection.map((item, position) => {
    const isLegacyIndex = typeof item === "number" || typeof item === "string" && /^\d+$/.test(item);
    const numericIndex = isLegacyIndex ? Number(item) : -1;
    const map = isLegacyIndex
      ? maps[numericIndex]
      : item && item.map !== undefined
        ? normalizeMapId_(item.map)
        : "";
    const index = maps.indexOf(map);

    if (index === -1 || seen.has(map)) return null;

    const requestedType = item && typeof item === "object"
      ? String(item.type || item.format || "")
      : "";
    const nita = item && typeof item === "object" && item.nita !== undefined
      ? Boolean(item.nita)
      : requestedType === "NI" || legacyNita.has(numericIndex);

    seen.add(map);
    return {
      map,
      index,
      type: nita ? "NI" : requestedType === "NI" ? "NI" : "Shroom",
      nita,
      position
    };
  }).filter(Boolean);
}
