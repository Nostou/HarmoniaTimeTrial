function TIME_TO_MS(t) {
  if (typeof t !== "string") return null;

  t = t.split("\n")[0].trim();

  const colon = t.indexOf(":");
  const dot = t.indexOf(".");

  if (colon === -1 || dot === -1) return null;

  const minutes = Number(t.slice(0, colon));
  const seconds = Number(t.slice(colon + 1, dot));
  const millis  = Number(t.slice(dot + 1));

  if (isNaN(minutes) || isNaN(seconds) || isNaN(millis)) return null;

  return (minutes * 60000 + seconds * 1000 + millis);
}

function MS_TO_TIME(ms) {
  if (ms === null || ms === "" || isNaN(ms)) return "";

  ms = Math.floor(ms);

  const minutes = Math.floor(ms / 60000);
  const seconds = Math.floor((ms % 60000) / 1000);
  const millis  = ms % 1000;

  return ( minutes + ":" +String(seconds).padStart(2, "0") + "." +String(millis).padStart(3, "0"));
}

function SUM_TIMES(range) {
  let total = 0;
  let validCount = 0;
  const flat = range.flat();

  flat.forEach(t => {
    const ms = TIME_TO_MS(t);
    if (ms !== null) {
      total += ms;
      validCount++;
    }
  });
  
  if (validCount === flat.length) return MS_TO_TIME(total);
  return `${validCount}/${flat.length}`;
}

function MULTIPLY_TIME(t, factor) {
  if (!t) return "-";

  let ms = TIME_TO_MS(t);
  if (!ms) return "-";

  ms *= factor;
  return MS_TO_TIME(ms);
}