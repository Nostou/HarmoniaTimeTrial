function TracksMarkdown() {
  const html = HtmlService.createHtmlOutputFromFile("TracksMarkdownPopup")
    .setWidth(300)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Sélectionner les courses");
}

function CreateTracksMarkdown(selectedIndices, selectedNita) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Archives");
  
  const headers = sheet.getRange("D1:M1").getDisplayValues()[0];
  const dataShrooms = sheet.getRange("B2:M31").getDisplayValues();
  const dataNita = sheet.getRange("B33:M62").getDisplayValues();

  const shrooms = selectedIndices.filter(i => !selectedNita.includes(i));
  const nita = selectedIndices.filter(i => selectedNita.includes(i));

  function formatLigne(row) {
    const cours = row[0];
    const moyenne = row[1];
    const temps = row.slice(2);

    let bestMs = Infinity;
    let bestTime = "";
    let bestJoueur = "";
    let count = 0;
    const total = temps.length;

    temps.forEach((t, i) => {
      const clean = t.split("\n")[0].trim();
      if (!clean || clean === "-") return;
      count++;
      const ms = TIME_TO_MS(clean);
      if (ms !== null && ms < bestMs) {
        bestMs = ms;
        bestTime = clean;
        bestJoueur = headers[i];
      }
    });

    const precision = `${count}/${total}`;
    return `> [${cours}]  [${precision}]\nMoyenne : ${moyenne.split("\n")[0].trim()}\nMeilleur : ${bestTime} (${bestJoueur})\n`;
  }

  let texte = `\`\`\`Gros TT n°XX - XX/XX/XXXX\`\`\`\n`;

  texte += `## Shrooms\n\`\`\`\n`;
  shrooms.forEach(i => {
    texte += formatLigne(dataShrooms[i]) + "\n";
  });
  texte += `\`\`\`\n`;

  if (nita.length > 0) {
    texte += `## NITA\n\`\`\`\n`;
    nita.forEach(i => {
      texte += formatLigne(dataNita[i]) + "\n";
    });
    texte += `\`\`\``;
  }

  return texte;
}