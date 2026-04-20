// Smoke test — port Slide C of handoff/build_showcase.js to the module.
// Success criterion: the rendered slide should match the reference Slide C
// pixel-close. If it doesn't, the module has the wrong abstraction.

const pptxgen = require("pptxgenjs");
const path = require("path");
const {
  saperia,
  PlotContext,
  areaScale,
  chrome,
} = require("../src");

const {
  bgFill,
  addEyebrow,
  addTitle,
  addSubtitle,
} = chrome;

// ── Data (from handoff/build_showcase.js) ─────────────────────────────
const years = [
  { year: 2025, util: 0.7133, rate: 247.22, rev: 24.09, phase: "setup" },
  { year: 2026, util: 0.6347, rate: 236.81, rev: 22.34, phase: "compression" },
  { year: 2027, util: 0.6235, rate: 235.47, rev: 22.13, phase: "compression" },
  { year: 2028, util: 0.6596, rate: 242.24, rev: 25.62, phase: "compression" },
  { year: 2029, util: 0.5958, rate: 247.75, rev: 25.29, phase: "compression" },
  { year: 2030, util: 0.7658, rate: 250.81, rev: 34.24, phase: "recovery" },
  { year: 2031, util: 0.7553, rate: 255.11, rev: 40.79, phase: "recovery" },
  { year: 2032, util: 0.7371, rate: 261.27, rev: 47.79, phase: "scale" },
  { year: 2033, util: 0.7053, rate: 268.44, rev: 51.07, phase: "scale" },
  { year: 2034, util: 0.7100, rate: 274.28, rev: 55.99, phase: "scale" },
  { year: 2035, util: 0.6755, rate: 282.98, rev: 58.66, phase: "scale" },
];

const phaseColors = {
  setup:       "MUTED",
  compression: "BERRY",
  recovery:    "SLATE",
  scale:       "STEEL",
};

// ── Build deck ─────────────────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title  = "claude-pptx-plot smoke test — Slide C port";
pres.author = "Saperia Consulting";

const t = saperia;
const s = pres.addSlide();

bgFill(s, t);
addEyebrow(s, t, "APPROACH C  ·  PURE NATIVE SHAPES — NO CHART OBJECT (via claude-pptx-plot)", t.layout.MARGIN, 0.5);
addTitle(s, t, "Every bubble is a shape.");
addSubtitle(
  s, t,
  "No chart object at all. Every bubble is an ellipse, the arc is a line, axes are lines. Full control over bubble sizing, colors, and labels. Editable as graphic design — not as a chart."
);

// ── Quadrant plot ──────────────────────────────────────────────────────
const plot = new PlotContext(s, {
  xRange: [0.55, 0.80],
  yRange: [225, 295],
  position: { x: 0.6, y: 1.85, size: 5.4 },
  padding: { left: 0.45, right: 0.1, top: 0.1, bottom: 0.4 },
  theme: t,
});

plot.quadrants({
  split: [0.70, 255],
  labels: [
    { corner: "tl", hdr: "LEVERAGED VOLUME", sub: "High utilization, modest rates",  tint: "LEVERAGED", transparency: 75 },
    { corner: "tr", hdr: "PREMIUM SCARCITY", sub: "High utilization, premium rates", tint: "PREMIUM",   transparency: 65 },
    { corner: "bl", hdr: "CRISIS",           sub: "Low utilization, low rates",      tint: "CRISIS",    transparency: 70 },
    { corner: "br", hdr: "MARGIN TRAP",      sub: "Falling util, high rates",        tint: "MARGIN",    transparency: 70 },
  ],
});

plot.frame();

plot.axes({
  x: {
    ticks: [0.55, 0.60, 0.65, 0.70, 0.75, 0.80],
    format: "pct",
    title: "UTILIZATION %",
  },
  y: {
    ticks: [230, 240, 250, 260, 270, 280, 290],
    format: "dollar",
    title: "AVG BILL RATE",
  },
});

// Arc: faint connecting line between consecutive years
plot.path(
  years.map((y) => ({ x: y.util, y: y.rate })),
  { color: "MUTED", width: 0.5, transparency: 55 }
);

// Bubbles — area-scaled, colored by phase
years.forEach((y) => {
  const dia = areaScale({ value: y.rev, domain: [22, 59], range: [0.22, 0.54] });
  plot.bubble({
    x: y.util,
    y: y.rate,
    size: dia,
    color: phaseColors[y.phase],
    transparency: 25,
    lineWidth: 1.25,
    label: y.year,
    labelColor: y.phase === "setup" ? "INK" : "WHITE",
  });
});

// ── Right-column insights (inline; no module helper yet — see notes) ──
{
  const rightX = plot.outer.x + plot.outer.w + 0.5;
  const rightW = t.layout.SLIDE_W - rightX - t.layout.MARGIN;
  const rightY = plot.outer.y + 0.1;

  s.addText([{ text: "2 years", options: { color: t.colors.INK, highlight: t.colors.LIME } }], {
    x: rightX, y: rightY, w: rightW, h: 0.9,
    fontFace: t.fonts.DISPLAY, fontSize: 44, margin: 0, valign: "top",
  });

  s.addText(
    [
      { text: "Crisis ",           options: { color: t.colors.BERRY, italic: true } },
      { text: "→",                 options: { color: t.colors.MUTED } },
      { text: " Premium Scarcity", options: { color: t.colors.INK, italic: true } },
    ],
    {
      x: rightX, y: rightY + 0.85, w: rightW, h: 0.4,
      fontFace: t.fonts.DISPLAY, fontSize: 15, margin: 0, valign: "top",
    }
  );

  s.addText(
    "Four years to fall, two to climb back. From 2029 (util 60%, rate $248) to 2031 (util 76%, rate $255).",
    {
      x: rightX, y: rightY + 1.35, w: rightW, h: 0.95,
      fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, color: t.colors.MUTED,
      margin: 0, valign: "top", lineSpacing: 15,
    }
  );

  chrome.addHairline(s, t, rightY + 2.4, { x: rightX, w: rightW });

  s.addText("2035 WARNING", {
    x: rightX, y: rightY + 2.55, w: rightW, h: 0.22,
    fontFace: t.fonts.SANS, fontSize: 10, bold: true, color: t.colors.MUTED,
    charSpacing: 2, valign: "top", margin: 0,
  });
  chrome.addInkUnderscore(s, t, rightX, rightY + 2.82, rightW * 0.25);

  s.addText([{ text: "−9 pts", options: { color: t.colors.BERRY } }], {
    x: rightX, y: rightY + 2.9, w: rightW, h: 0.75,
    fontFace: t.fonts.DISPLAY, fontSize: 34, margin: 0, valign: "top",
  });

  s.addText("Utilization slipping; rates still climbing. Classic Margin Trap drift.", {
    x: rightX, y: rightY + 3.7, w: rightW, h: 0.8,
    fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, color: t.colors.MUTED,
    margin: 0, valign: "top", lineSpacing: 15,
  });
}

// ── Export ─────────────────────────────────────────────────────────────
const outPath = path.resolve(__dirname, "..", "output", "smoke_slide_c.pptx");
pres.writeFile({ fileName: outPath })
  .then((f) => console.log("Wrote:", f))
  .catch((err) => { console.error(err); process.exit(1); });
