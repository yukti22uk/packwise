// ─── UTILITY FUNCTIONS ───────────────────────────────────────────────────────
function fmtN(v){return Number.isInteger(v)?String(v):parseFloat(v.toFixed(1)).toString();}
function money(v){return v.toLocaleString(undefined,{maximumFractionDigits:2});}


export { fmtN, money };
