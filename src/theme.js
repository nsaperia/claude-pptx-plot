// Theme — Saperia design system.
// Hex values are pptxgenjs-shaped: no `#` prefix.
// Lifted from handoff/build_showcase.js and handoff/build_deck.js.
// To swap brands: copy saperia, change tokens, pass to ClaudePPTXPlot({ theme }).

const saperia = {
  name: "saperia",

  colors: {
    BG:        "FFF5ED",
    BG_RAISED: "FBEFE4",
    INK:       "353745",
    MUTED:     "8A7968",
    RULE:      "C9B9A8",
    LIME:      "C1FF72",
    BERRY:     "B84C65",
    STEEL:     "2D5F7C",
    SLATE:     "4A5E6A",
    LBLUE:     "BFE5EF",
    GOLD:      "C5A55A",
    LIGHT:     "FBFCFC",
    WHITE:     "FFFFFF",
  },

  // Quadrant background tints — lifted verbatim from build_showcase.js.
  quadrantTints: {
    LEVERAGED: "E9DED0",
    PREMIUM:   "C1FF72",  // LIME
    CRISIS:    "E6C9BF",
    MARGIN:    "D4DEE5",
  },

  fonts: {
    DISPLAY: "Georgia",
    BODY:    "Georgia",
    SANS:    "Calibri",
  },

  layout: {
    SLIDE_W:   13.33,
    SLIDE_H:   7.5,
    MARGIN:    0.6,
    LOGO_W:    1.6,
    LOGO_H:    1.6 * (1764 / 3000),
  },

  // Data-series color policies. Enforces the "never STEEL + SLATE" rule.
  seriesColors: {
    single: ["STEEL"],
    pair:   ["STEEL", "BERRY"],
    triple: ["STEEL", "LBLUE", "BERRY"],  // LBLUE gets heavier stroke at render time
  },

  // Stroke weights — LBLUE reads poorly at default 2.5pt on BG.
  strokeForColor: (colorName) => (colorName === "LBLUE" ? 3.5 : 2.5),
  symbolSizeForColor: (colorName) => (colorName === "LBLUE" ? 8 : 6),
};

module.exports = { saperia };
