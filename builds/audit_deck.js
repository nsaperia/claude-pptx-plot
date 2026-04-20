// Chart-type audit deck — Round 1.
// Two chart types × three approaches = six slides.
// Column (native wins easy) + Waterfall (pure shapes wins).
//
// This deck's job is NOT to look perfect — it's to surface, side by side,
// the visual tradeoffs between native pptxgenjs charts and shape-based
// rendering via claude-pptx-plot.

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { saperia, PlotContext, chrome } = require("../src");
const {
  bgFill, addEyebrow, addTitle, addSubtitle,
  addHairline, addInkUnderscore,
} = chrome;

// ── Data ──────────────────────────────────────────────────────────────
const deckData = JSON.parse(
  fs.readFileSync(path.resolve(__dirname, "..", "..", "handoff", "deck_data.json"), "utf8")
);

// Revenue in $M for Column chart
const years = deckData.annual.map((r) => r.year);
const revenueM = deckData.annual.map((r) => +(r.revenue / 1e6).toFixed(2));

// EBITDA $ walk for Waterfall: 2025 → 2030 (covers the trough-and-peak story)
const walkYears = [2025, 2026, 2027, 2028, 2029, 2030];
const walkEbitdaM = walkYears.map((y) => {
  const row = deckData.annual.find((r) => r.year === y);
  return +(row.ebitdaPct * row.revenue / 1e6).toFixed(2);
});
// walkEbitdaM: [6.53, 4.69, 4.32, 6.19, 3.87, 11.43]

// ── Presentation ──────────────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title  = "Chart-type audit deck — Column + Waterfall";
pres.author = "Saperia Consulting";

const t = saperia;
const { SLIDE_W, SLIDE_H, MARGIN } = t.layout;

function header(s, eyebrow, title, subtitle) {
  bgFill(s, t);
  addEyebrow(s, t, eyebrow, MARGIN, 0.5);
  addTitle(s, t, title);
  if (subtitle) addSubtitle(s, t, subtitle);
}

// Column slide layout: wide plot, no right column
const COL_POS = { x: 0.7, y: 2.2, w: SLIDE_W - 1.4, h: 4.4 };
const COL_PAD = { left: 0.6, right: 0.15, top: 0.1, bottom: 0.4 };
const COL_X_RANGE = [-0.5, years.length - 0.5];
const COL_Y_RANGE = [0, 65];
const COL_Y_TICKS = [0, 10, 20, 30, 40, 50, 60];

// Waterfall slide layout: same frame, narrower x-range (7 bars)
const WF_POS = { x: 0.7, y: 2.2, w: SLIDE_W - 1.4, h: 4.4 };
const WF_PAD = { left: 0.6, right: 0.15, top: 0.1, bottom: 0.4 };
const WF_X_RANGE = [-0.5, walkYears.length + 0.5];  // 7 slots: anchor + 5 floats + anchor
const WF_Y_RANGE = [0, 13];
const WF_Y_TICKS = [0, 2, 4, 6, 8, 10, 12];

// Derive the walk: each entry = { kind, x, value, baseline, delta }
const walk = [];
// 2025 start anchor
walk.push({ kind: "anchor", year: 2025, x: 0, value: walkEbitdaM[0], baseline: 0, delta: null });
// Floats 2026..2030
for (let i = 1; i < walkEbitdaM.length; i++) {
  const prev = walkEbitdaM[i - 1];
  const cur  = walkEbitdaM[i];
  walk.push({
    kind: cur >= prev ? "pos" : "neg",
    year: walkYears[i],
    x: i,
    value: cur,
    baseline: prev,
    delta: +(cur - prev).toFixed(2),
  });
}
// 2030 end anchor
walk.push({
  kind: "anchor", year: 2030, x: walkYears.length, value: walkEbitdaM[walkEbitdaM.length - 1], baseline: 0, delta: null,
});

// ══════════════════════════════════════════════════════════════════════
// SLIDE 1A — Column, pure native
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "APPROACH A  ·  PURE NATIVE PPT COLUMN CHART",
    "One chart, nothing else.",
    "Pure pptxgenjs addChart. Saperia palette and fonts passed in; everything else is PowerPoint defaults."
  );

  s.addChart(
    pres.ChartType.bar,
    [{ name: "Revenue ($M)", labels: years, values: revenueM }],
    {
      x: COL_POS.x, y: COL_POS.y, w: COL_POS.w, h: COL_POS.h,
      barDir: "col",
      chartColors: [t.colors.STEEL],
      showLegend: false,
      showValAxisTitle: true, valAxisTitle: "Revenue ($M)",
      valAxisTitleFontFace: t.fonts.SANS, valAxisTitleFontSize: 10, valAxisTitleColor: t.colors.MUTED,
      valAxisLabelFormatCode: "$0\"M\"",
      catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
      valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
      valGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
      plotArea:  { fill: { color: t.colors.BG } },
      chartArea: { fill: { color: t.colors.BG } },
      barGapWidthPct: 55,
    }
  );
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 1B — Column, native + overlay chrome
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "APPROACH B  ·  NATIVE COLUMN CHART + OVERLAY CHROME",
    "Native chart, shapes on top.",
    "Native chart for the plot area. Axis title, hero callout, and value labels live as overlay shapes/text."
  );

  // Native chart — minimal built-in chrome; overlay does the work
  s.addChart(
    pres.ChartType.bar,
    [{ name: "Revenue ($M)", labels: years, values: revenueM }],
    {
      x: COL_POS.x, y: COL_POS.y, w: COL_POS.w, h: COL_POS.h,
      barDir: "col",
      chartColors: [t.colors.STEEL],
      showLegend: false,
      showValAxisTitle: false,
      valAxisLabelFormatCode: "$0\"M\"",
      catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
      valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
      valGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
      plotArea:  { fill: { color: t.colors.BG } },
      chartArea: { fill: { color: t.colors.BG } },
      barGapWidthPct: 55,
    }
  );

  // Overlay: hero callout highlighting 2029 trough and 2030 recovery
  s.addText([
    { text: "2029 → 2030: ", options: { color: t.colors.INK, bold: true } },
    { text: "+35%", options: { color: t.colors.INK, highlight: t.colors.LIME, bold: true } },
    { text: " in one year", options: { color: t.colors.INK } },
  ], {
    x: MARGIN, y: 6.7, w: SLIDE_W - 2 * MARGIN, h: 0.35,
    fontFace: t.fonts.DISPLAY, fontSize: 13, italic: true, valign: "top", margin: 0,
  });

  // A thin overlay rule under the callout to tie it visually to the chart
  addHairline(s, t, 6.65);
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 1C — Column, pure shapes via claude-pptx-plot
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "APPROACH C  ·  PURE NATIVE SHAPES — NO CHART OBJECT",
    "Every bar is a rect.",
    "No chart object. Rectangles at computed positions, axis ticks as text runs. More code; full editorial control."
  );

  const plot = new PlotContext(s, {
    xRange: COL_X_RANGE,
    yRange: COL_Y_RANGE,
    position: COL_POS,
    padding: COL_PAD,
    theme: t,
  });

  plot.frame();
  plot.axes({
    x: {
      ticks: years.map((_, i) => i),
      format: (i) => String(years[i]),
      title: "YEAR",
    },
    y: {
      ticks: COL_Y_TICKS,
      format: (v) => `$${v}M`,
      title: "REVENUE ($M)",
    },
  });

  // Horizontal gridlines at each y-tick (behind bars)
  COL_Y_TICKS.slice(1).forEach((v) => {
    const y = plot.yToSlide(v);
    const r = plot.plotRect();
    s.addShape("line", {
      x: r.x, y, w: r.w, h: 0,
      line: { color: t.colors.RULE, width: 0.4, transparency: 40 },
    });
  });

  // Bars
  revenueM.forEach((rev, i) => {
    plot.bar({
      x: i,
      value: rev,
      width: 0.65,
      color: "STEEL",
      transparency: 0,
    });
  });

  // One LIME hero: label the 2030 jump
  const i2030 = years.indexOf(2030);
  if (i2030 >= 0) {
    const cx = plot.xToSlide(i2030);
    const ty = plot.yToSlide(revenueM[i2030]);
    s.addText("2030", {
      x: cx - 0.5, y: ty - 0.42, w: 1.0, h: 0.3,
      fontFace: t.fonts.SANS, fontSize: 9, bold: true,
      color: t.colors.INK, highlight: t.colors.LIME,
      align: "center", valign: "bottom", margin: 0, charSpacing: 1,
    });
  }
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 4A — Waterfall, native stacked-column hack
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "APPROACH A  ·  NATIVE STACKED-COLUMN HACK",
    "Not a waterfall. The closest native can get.",
    "pptxgenjs has no waterfall type. Best-effort: stacked column with an invisible base series. No connectors. One color for every delta — positives indistinguishable from negatives."
  );

  // Base series (invisible, BG-colored) + delta series (single color)
  const baseSeries = walk.map((w) =>
    w.kind === "anchor" ? 0 : Math.min(w.baseline, w.value)
  );
  const deltaSeries = walk.map((w) =>
    w.kind === "anchor" ? w.value : Math.abs(w.value - w.baseline)
  );
  const catLabels = walk.map((w) => (w.kind === "anchor" ? `${w.year}` : `${w.year}`));

  s.addChart(
    pres.ChartType.bar,
    [
      { name: "Base",  labels: catLabels, values: baseSeries },
      { name: "Delta", labels: catLabels, values: deltaSeries },
    ],
    {
      x: WF_POS.x, y: WF_POS.y, w: WF_POS.w, h: WF_POS.h,
      barDir: "col",
      barGrouping: "stacked",
      chartColors: [t.colors.BG, t.colors.STEEL],   // base blends into BG
      showLegend: false,
      showValAxisTitle: true, valAxisTitle: "EBITDA ($M)",
      valAxisTitleFontFace: t.fonts.SANS, valAxisTitleFontSize: 10, valAxisTitleColor: t.colors.MUTED,
      valAxisLabelFormatCode: "$0\"M\"",
      valAxisMinVal: 0, valAxisMaxVal: 13,
      catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
      valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
      valGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
      plotArea:  { fill: { color: t.colors.BG } },
      chartArea: { fill: { color: t.colors.BG } },
      barGapWidthPct: 40,
    }
  );
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 4B — Waterfall, native hack + overlay chrome
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "APPROACH B  ·  NATIVE STACKED-COLUMN HACK + OVERLAY",
    "Labels rescued, colors cannot be.",
    "Same native hack. Overlay adds delta labels and step connectors. Bars themselves are still one color — pptxgenjs does not expose per-bar color in native charts."
  );

  const baseSeries = walk.map((w) =>
    w.kind === "anchor" ? 0 : Math.min(w.baseline, w.value)
  );
  const deltaSeries = walk.map((w) =>
    w.kind === "anchor" ? w.value : Math.abs(w.value - w.baseline)
  );
  const catLabels = walk.map((w) => `${w.year}`);

  s.addChart(
    pres.ChartType.bar,
    [
      { name: "Base",  labels: catLabels, values: baseSeries },
      { name: "Delta", labels: catLabels, values: deltaSeries },
    ],
    {
      x: WF_POS.x, y: WF_POS.y, w: WF_POS.w, h: WF_POS.h,
      barDir: "col",
      barGrouping: "stacked",
      chartColors: [t.colors.BG, t.colors.STEEL],
      showLegend: false,
      showValAxisTitle: true, valAxisTitle: "EBITDA ($M)",
      valAxisTitleFontFace: t.fonts.SANS, valAxisTitleFontSize: 10, valAxisTitleColor: t.colors.MUTED,
      valAxisLabelFormatCode: "$0\"M\"",
      valAxisMinVal: 0, valAxisMaxVal: 13,
      catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
      valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
      valGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
      plotArea:  { fill: { color: t.colors.BG } },
      chartArea: { fill: { color: t.colors.BG } },
      barGapWidthPct: 40,
    }
  );

  // Overlay: build a PlotContext whose x/y align to the native chart's
  // category axis. pptxgenjs category charts put the first category center
  // roughly at plot_left + (catWidth/2). Matching this exactly is the
  // fragile part of overlay approaches — if the chart resizes, overlays drift.
  const plot = new PlotContext(s, {
    xRange: WF_X_RANGE,
    yRange: WF_Y_RANGE,
    position: WF_POS,
    padding: WF_PAD,
    theme: t,
  });

  // Horizontal connectors between adjacent bar tops (approximate x-positions)
  for (let i = 0; i < walk.length - 1; i++) {
    const a = walk[i], b = walk[i + 1];
    const ax = plot.xToSlide(a.x + 0.32);    // right edge of bar a
    const bx = plot.xToSlide(b.x - 0.32);    // left edge of bar b
    const ay = plot.yToSlide(a.value);
    s.addShape("line", {
      x: ax, y: ay, w: bx - ax, h: 0,
      line: { color: t.colors.MUTED, width: 0.75, dashType: "dash", transparency: 30 },
    });
  }

  // Delta labels — one per float bar, above the bar top
  walk.filter((w) => w.kind !== "anchor").forEach((w) => {
    const cx = plot.xToSlide(w.x);
    const topV = Math.max(w.baseline, w.value);
    const cy = plot.yToSlide(topV);
    const sign = w.delta >= 0 ? "+" : "−";
    const abs = Math.abs(w.delta).toFixed(1);
    s.addText(`${sign}$${abs}M`, {
      x: cx - 0.5, y: cy - 0.35, w: 1.0, h: 0.28,
      fontFace: t.fonts.SANS, fontSize: 9, bold: true,
      color: w.delta >= 0 ? t.colors.STEEL : t.colors.BERRY,
      align: "center", valign: "bottom", margin: 0,
    });
  });
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 4C — Waterfall, pure shapes via claude-pptx-plot
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "APPROACH C  ·  PURE NATIVE SHAPES — NO CHART OBJECT",
    "A real waterfall.",
    "Each bar is a rect. Connectors are lines. Positive steps in STEEL, negative in BERRY, anchors in INK. Delta labels above each float; anchors labeled with totals."
  );

  const plot = new PlotContext(s, {
    xRange: WF_X_RANGE,
    yRange: WF_Y_RANGE,
    position: WF_POS,
    padding: WF_PAD,
    theme: t,
  });

  plot.frame();
  plot.axes({
    x: {
      ticks: walk.map((w) => w.x),
      format: (i) => String(walk[i]?.year ?? ""),
      title: "YEAR",
    },
    y: {
      ticks: WF_Y_TICKS,
      format: (v) => `$${v}M`,
      title: "EBITDA ($M)",
    },
  });

  // Gridlines
  WF_Y_TICKS.slice(1).forEach((v) => {
    const y = plot.yToSlide(v);
    const r = plot.plotRect();
    s.addShape("line", {
      x: r.x, y, w: r.w, h: 0,
      line: { color: t.colors.RULE, width: 0.4, transparency: 40 },
    });
  });

  // Connectors (drawn first so bars sit on top)
  for (let i = 0; i < walk.length - 1; i++) {
    const a = walk[i], b = walk[i + 1];
    const ax = plot.xToSlide(a.x + 0.32);
    const bx = plot.xToSlide(b.x - 0.32);
    const ay = plot.yToSlide(a.value);
    s.addShape("line", {
      x: ax, y: ay, w: bx - ax, h: 0,
      line: { color: t.colors.MUTED, width: 0.75, dashType: "dash", transparency: 30 },
    });
  }

  // Bars
  walk.forEach((w) => {
    const color =
      w.kind === "anchor" ? "INK" :
      w.kind === "pos"    ? "STEEL" :
                            "BERRY";
    plot.bar({
      x: w.x,
      value: w.value,
      baseline: w.kind === "anchor" ? 0 : w.baseline,
      width: 0.65,
      color,
      transparency: w.kind === "anchor" ? 15 : 0,
    });

    // Labels: anchors show total; floats show delta
    const cx = plot.xToSlide(w.x);
    if (w.kind === "anchor") {
      const cy = plot.yToSlide(w.value);
      s.addText(`$${w.value.toFixed(1)}M`, {
        x: cx - 0.6, y: cy - 0.38, w: 1.2, h: 0.3,
        fontFace: t.fonts.SANS, fontSize: 10, bold: true,
        color: t.colors.INK,
        align: "center", valign: "bottom", margin: 0,
      });
    } else {
      const topV = Math.max(w.baseline, w.value);
      const cy = plot.yToSlide(topV);
      const sign = w.delta >= 0 ? "+" : "−";
      const abs = Math.abs(w.delta).toFixed(1);
      s.addText(`${sign}$${abs}M`, {
        x: cx - 0.6, y: cy - 0.38, w: 1.2, h: 0.3,
        fontFace: t.fonts.SANS, fontSize: 10, bold: true,
        color: w.delta >= 0 ? t.colors.STEEL : t.colors.BERRY,
        align: "center", valign: "bottom", margin: 0,
      });
    }
  });

  // Hero callout tying the slide to the story
  s.addText([
    { text: "Four years of compression, then ",         options: { color: t.colors.INK } },
    { text: "+$7.6M in one year",                        options: { color: t.colors.INK, highlight: t.colors.LIME, bold: true } },
    { text: " of EBITDA snap-back.",                     options: { color: t.colors.INK } },
  ], {
    x: MARGIN, y: 6.75, w: SLIDE_W - 2 * MARGIN, h: 0.32,
    fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, valign: "top", margin: 0,
  });
}

// ── Write ─────────────────────────────────────────────────────────────
const outPath = path.resolve(__dirname, "..", "output", "audit_deck.pptx");
pres.writeFile({ fileName: outPath })
  .then((f) => console.log("Wrote:", f))
  .catch((err) => { console.error(err); process.exit(1); });
