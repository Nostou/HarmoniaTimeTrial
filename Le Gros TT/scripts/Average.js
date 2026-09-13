function readAveragesFromSheet(sheet, maps) {
  const averages = {};

  maps.forEach(m => {
    averages[m] = { Shroom: null, NI: null };
  });

  // --- SHROOM
  const shroomRange = sheet
    .getRange(2, 3, maps.length, 1)
    .getDisplayValues();

  shroomRange.forEach((row, i) => {
    const v = TIME_TO_MS(row[0]);
    if (v !== null) {
      averages[maps[i]].Shroom = v;
    }
  });

  // --- NI
  const niStartRow = maps.length + 3;
  const niRange = sheet
    .getRange(niStartRow, 3, maps.length, 1)
    .getDisplayValues();

  niRange.forEach((row, i) => {
    const v = TIME_TO_MS(row[0]);
    if (v !== null) {
      averages[maps[i]].NI = v;
    }
  });

  return averages;
}

function computeAverageResults(results, maps, players, averageKey = "Average") {
  maps.forEach(map => {
    results[map].__avgCount = { Shroom: 0, NI: 0 };

    ["Shroom", "NI"].forEach(type => {
      let sum = 0;
      let count = 0;

      players.forEach(player => {
        const t = results[map][player][type];
        if (t !== 9999999) {
          sum += t;
          count++;
        }
      });

      if (count > 0) {
        results[map][averageKey][type] = Math.round(sum / count);
        results[map].__avgCount[type] = count;
      }
    });
  });
}

function logAverageChanges(oldAvg, newResults, maps, averageKey = "Average") {
  const total = getPlayers().length;

  maps.forEach(map => {
    ["Shroom", "NI"].forEach(type => {
      const before = oldAvg[map]?.[type];
      const after  = newResults[map]?.[averageKey]?.[type];
      const count  = newResults[map].__avgCount?.[type] ?? 0;

      if (before == null) return;
      if (after == null) return;
      if (before === after) return;

      logUI(
        `${map} [${type}] : ${MS_TO_TIME(before)} → ${MS_TO_TIME(after)} [${count}/${total}]`
      );
    });
  });
}