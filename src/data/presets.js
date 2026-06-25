// ─── VEHICLE & PALLET PRESETS ────────────────────────────────────────────────
const PALLET_BASES = [
  { label:"1200×1000 mm", L:1200, W:1000 },
  { label:"1200×1200 mm", L:1200, W:1200 },
  { label:"1200×800 mm",  L:1200, W:800  },
];
const VEHICLES = [
  { label:"Tata Ace",             L:2100,  W:1525, H:1600 },
  { label:"19ft",                 L:5800,  W:2350, H:2100 },
  { label:"20ft Container (ISO)", L:5900,  W:2350, H:2390 },
  { label:"22ft",                 L:6700,  W:2350, H:2100 },
  { label:"32ft SXL",             L:9750,  W:2350, H:2700 },
  { label:"32ft MXL",             L:9750,  W:2430, H:2900 },
  { label:"40ft Container (ISO)", L:12032, W:2352, H:2395 },
];
const VEHICLES_WITH_CUSTOM=[...VEHICLES,{label:"Custom (Manual Input)",L:null,W:null,H:null}];


export { PALLET_BASES, VEHICLES, VEHICLES_WITH_CUSTOM };
