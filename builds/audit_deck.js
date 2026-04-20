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

  plot.grid({ y: COL_Y_TICKS });

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

  plot.grid({ y: WF_Y_TICKS });

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

// ══════════════════════════════════════════════════════════════════════
//
//                      ROUND 2 · NATIVE-ONLY SLIDES
//
//  Four additional chart types, one native slide each. Following the
//  Column + Waterfall round, user verdict locked in native as the default
//  and shape-based as the escape hatch. These four chart types show
//  native handling the job — no A/B/C comparison needed.
//
// ══════════════════════════════════════════════════════════════════════

// Shared content region for Line / Scatter (wide) and Pie / Radar (split).
const WIDE_POS  = { x: 0.7, y: 2.2, w: SLIDE_W - 1.4, h: 4.4 };
const SPLIT_LEFT  = { x: 0.7,  y: 2.2, w: 6.5, h: 4.4 };
const SPLIT_RIGHT = { x: 7.4,  y: 2.3, w: 5.3, h: 4.3 };

// ══════════════════════════════════════════════════════════════════════
// SLIDE 7 — Line, native
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "LINE CHART  ·  NATIVE",
    "EBITDA % compressed, then snapped back.",
    "Native pptxgenjs line chart. Saperia palette, axes, gridlines. Callout highlights the trough. No overlay — native handles this type cleanly."
  );

  s.addChart(
    pres.ChartType.line,
    [{
      name: "EBITDA %",
      labels: years,
      values: deckData.annual.map((r) => +(r.ebitdaPct * 100).toFixed(1)),
    }],
    {
      x: WIDE_POS.x, y: WIDE_POS.y, w: WIDE_POS.w, h: WIDE_POS.h,
      chartColors: [t.colors.STEEL],
      lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      showLegend: false,
      showValAxisTitle: true, valAxisTitle: "EBITDA %",
      valAxisTitleFontFace: t.fonts.SANS, valAxisTitleFontSize: 10, valAxisTitleColor: t.colors.MUTED,
      valAxisLabelFormatCode: "0\"%\"",
      valAxisMinVal: 10, valAxisMaxVal: 35,
      catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
      valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
      valGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
      plotArea:  { fill: { color: t.colors.BG } },
      chartArea: { fill: { color: t.colors.BG } },
    }
  );

  s.addText([
    { text: "2029 trough: ",                options: { color: t.colors.INK } },
    { text: "15.3%",                         options: { color: t.colors.INK, highlight: t.colors.LIME, bold: true } },
    { text: " — a 12-point compression from 2025. Back to 33% by 2030.", options: { color: t.colors.INK } },
  ], {
    x: MARGIN, y: 6.75, w: SLIDE_W - 2 * MARGIN, h: 0.32,
    fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, valign: "top", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 8 — Pie / Donut, native
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "DONUT CHART  ·  NATIVE",
    "Three practices, split close to thirds by 2035.",
    "Native pptxgenjs doughnut. Text callouts on the right replace native labels (the common native weakness)."
  );

  // 2035 revenue by practice ($M)
  const mix2035 = {
    Strategy:   +(deckData.practiceYear.Strategy[10]   / 1e6).toFixed(2),
    Operations: +(deckData.practiceYear.Operations[10] / 1e6).toFixed(2),
    Technology: +(deckData.practiceYear.Technology[10] / 1e6).toFixed(2),
  };
  const total2035 = +(mix2035.Strategy + mix2035.Operations + mix2035.Technology).toFixed(2);

  s.addChart(
    pres.ChartType.doughnut,
    [{
      name: "Revenue mix",
      labels: ["Strategy", "Operations", "Technology"],
      values: [mix2035.Strategy, mix2035.Operations, mix2035.Technology],
    }],
    {
      x: SPLIT_LEFT.x, y: SPLIT_LEFT.y, w: SPLIT_LEFT.w, h: SPLIT_LEFT.h,
      chartColors: [t.colors.STEEL, t.colors.LBLUE, t.colors.BERRY],
      showLegend: false,
      // Native data labels turned OFF — right column carries the readout.
      // Native donut labels sit inside the slice at a fixed size and are
      // hard to style against a warm background; offloading to text
      // blocks gives proper typographic control.
      showValue: false, showPercent: false, showCategoryName: false,
      holeSize: 65,
      plotArea:  { fill: { color: t.colors.BG } },
      chartArea: { fill: { color: t.colors.BG } },
    }
  );

  // Right column — textual callouts replace native legend
  const rx = SPLIT_RIGHT.x;
  const ry = SPLIT_RIGHT.y;
  const rw = SPLIT_RIGHT.w;

  addEyebrow(s, t, "2035 TOTAL", rx, ry, rw);
  addInkUnderscore(s, t, rx, ry + 0.3, 1.5);
  s.addText(`$${total2035.toFixed(1)}M`, {
    x: rx, y: ry + 0.4, w: rw, h: 0.9,
    fontFace: t.fonts.DISPLAY, fontSize: 44, color: t.colors.INK, margin: 0, valign: "top",
  });

  const rows = [
    { name: "Operations", value: mix2035.Operations, color: t.colors.LBLUE, note: "Leader by 2035" },
    { name: "Technology", value: mix2035.Technology, color: t.colors.BERRY, note: "Fastest grower" },
    { name: "Strategy",   value: mix2035.Strategy,   color: t.colors.STEEL, note: "Cyclical anchor" },
  ];
  let y = ry + 1.6;
  rows.forEach((r) => {
    s.addShape("rect", { x: rx, y: y + 0.08, w: 0.14, h: 0.14, fill: { color: r.color }, line: { type: "none" } });
    s.addText(r.name, {
      x: rx + 0.3, y, w: rw - 0.3, h: 0.28,
      fontFace: t.fonts.SANS, fontSize: 11, bold: true, color: t.colors.INK,
      charSpacing: 1, valign: "top", margin: 0,
    });
    s.addText(`$${r.value.toFixed(1)}M  ·  ${((r.value / total2035) * 100).toFixed(0)}%`, {
      x: rx + 0.3, y: y + 0.28, w: rw - 0.3, h: 0.28,
      fontFace: t.fonts.DISPLAY, fontSize: 12, color: t.colors.INK, margin: 0, valign: "top",
    });
    s.addText(r.note, {
      x: rx + 0.3, y: y + 0.55, w: rw - 0.3, h: 0.3,
      fontFace: t.fonts.DISPLAY, fontSize: 11, italic: true, color: t.colors.MUTED,
      margin: 0, valign: "top",
    });
    y += 1.0;
  });
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 9 — Scatter (non-bubble), native
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "SCATTER PLOT  ·  NATIVE",
    "Utilization predicts EBITDA margin.",
    "Native pptxgenjs scatter chart (non-bubble). 11 years of (util, EBITDA %) points. Native handles scatter when sizing uniformity is fine."
  );

  // pptxgenjs scatter format: [{ name:"X", values:[...] }, { name:"Y1", values:[...] }]
  const scatterData = [
    { name: "Utilization %",  values: deckData.annual.map((r) => +(r.util * 100).toFixed(1)) },
    { name: "EBITDA %",       values: deckData.annual.map((r) => +(r.ebitdaPct * 100).toFixed(1)) },
  ];

  s.addChart(pres.ChartType.scatter, scatterData, {
    x: WIDE_POS.x, y: WIDE_POS.y, w: WIDE_POS.w, h: WIDE_POS.h,
    chartColors: [t.colors.STEEL],
    lineSize: 0,                          // scatter = dots only, no connecting line
    lineDataSymbol: "circle", lineDataSymbolSize: 9,
    lineDataSymbolLineColor: t.colors.STEEL, lineDataSymbolLineSize: 1,
    showLegend: false,
    showValAxisTitle: true, valAxisTitle: "EBITDA %",
    showCatAxisTitle: true, catAxisTitle: "Utilization %",
    valAxisTitleFontFace: t.fonts.SANS, valAxisTitleFontSize: 10, valAxisTitleColor: t.colors.MUTED,
    catAxisTitleFontFace: t.fonts.SANS, catAxisTitleFontSize: 10, catAxisTitleColor: t.colors.MUTED,
    valAxisLabelFormatCode: "0\"%\"",
    catAxisLabelFormatCode: "0\"%\"",
    valAxisMinVal: 10, valAxisMaxVal: 35,
    catAxisMinVal: 55, catAxisMaxVal: 80,
    valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
    catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
    valGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
    catGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
    plotArea:  { fill: { color: t.colors.BG } },
    chartArea: { fill: { color: t.colors.BG } },
  });

  s.addText(
    "Positive slope: every 10 utilization points ≈ +7 points of EBITDA. Pricing alone never closed the gap.",
    {
      x: MARGIN, y: 6.75, w: SLIDE_W - 2 * MARGIN, h: 0.32,
      fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, color: t.colors.INK,
      valign: "top", margin: 0,
    }
  );
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 10 — Radar, native (simple)
// ══════════════════════════════════════════════════════════════════════
// Replaces the round-2 three-axis normalized radar, which degenerated to
// a triangle with one practice collapsing to a spike. Six time-points
// (every other year) gives a proper radar shape; $M values share a single
// scale so no normalization distortion.
{
  const s = pres.addSlide();
  header(s,
    "RADAR CHART  ·  NATIVE",
    "Three practices, three different trajectories.",
    "Six time-points on six axes (every other year, 2025 to 2035). Revenue in $M — one shared scale, no normalization. Native handles multi-series radar when axes share units."
  );

  const radarYears = ["2025", "2027", "2029", "2031", "2033", "2035"];
  const radarIdx   = [0, 2, 4, 6, 8, 10];   // indexes into practiceYear arrays
  const seriesAt = (name) => ({
    name,
    labels: radarYears,
    values: radarIdx.map((i) => +(deckData.practiceYear[name][i] / 1e6).toFixed(2)),
  });
  const radarSeries = [seriesAt("Strategy"), seriesAt("Operations"), seriesAt("Technology")];

  s.addChart(pres.ChartType.radar, radarSeries, {
    x: SPLIT_LEFT.x, y: SPLIT_LEFT.y, w: SPLIT_LEFT.w, h: SPLIT_LEFT.h,
    chartColors: [t.colors.STEEL, t.colors.LBLUE, t.colors.BERRY],
    radarStyle: "standard",
    lineSize: 2.5,
    showLegend: true, legendPos: "b",
    legendFontSize: 10, legendFontFace: t.fonts.SANS, legendColor: t.colors.MUTED,
    catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
    valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 9,  valAxisLabelColor: t.colors.MUTED,
    valAxisMinVal: 0, valAxisMaxVal: 22,
    valAxisLabelFormatCode: "$0\"M\"",
    valGridLine: { color: t.colors.RULE, style: "solid", size: 0.4 },
    plotArea:  { fill: { color: t.colors.BG } },
    chartArea: { fill: { color: t.colors.BG } },
  });

  // Right column — interpretation (updated for the time-radar narrative)
  const rx = SPLIT_RIGHT.x;
  const ry = SPLIT_RIGHT.y;
  const rw = SPLIT_RIGHT.w;

  addEyebrow(s, t, "READ THE SHAPES", rx, ry, rw);
  addInkUnderscore(s, t, rx, ry + 0.3, 1.5);

  const notes = [
    { color: t.colors.STEEL, name: "Strategy",   body: "Tightest shape. Shallow dip in 2027, steady climb after. Least cyclical." },
    { color: t.colors.LBLUE, name: "Operations", body: "Deepest dent in 2029, then the biggest snap-back — ends largest in 2035." },
    { color: t.colors.BERRY, name: "Technology", body: "Grew through every year, downturn included. The most counter-cyclical practice." },
  ];
  let y = ry + 0.95;
  notes.forEach((n) => {
    s.addShape("rect", { x: rx, y: y + 0.08, w: 0.14, h: 0.14, fill: { color: n.color }, line: { type: "none" } });
    s.addText(n.name, {
      x: rx + 0.3, y, w: rw - 0.3, h: 0.28,
      fontFace: t.fonts.SANS, fontSize: 11, bold: true, color: t.colors.INK,
      charSpacing: 1, valign: "top", margin: 0,
    });
    s.addText(n.body, {
      x: rx + 0.3, y: y + 0.28, w: rw - 0.3, h: 0.9,
      fontFace: t.fonts.DISPLAY, fontSize: 11, italic: true, color: t.colors.MUTED,
      margin: 0, valign: "top", lineSpacing: 14,
    });
    y += 1.15;
  });
}

// ══════════════════════════════════════════════════════════════════════
//
//                     ROUND 3 · STRESS-TEST SLIDES
//
// One stress-test per native chart type. Simple versions (above) show
// native handling the easy case; stress-tests show where native's
// ceiling is — combo charts, per-series styling, multi-point data,
// filled overlays. Together they give the scorecard forensic evidence
// at two complexity tiers.
//
// ══════════════════════════════════════════════════════════════════════

// ══════════════════════════════════════════════════════════════════════
// SLIDE 11 — Line, stress-test (dual-axis combo)
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "LINE CHART  ·  STRESS-TEST (DUAL-AXIS COMBO)",
    "Revenue bars, EBITDA % line, secondary axis.",
    "Multi-type chart: native pptxgenjs pres.ChartType has no 'combo' — you pass an array of typed blocks. Each block has its own options. Tests secondary axis, per-series styling, legend mixing."
  );

  const comboData = [
    {
      type: pres.ChartType.bar,
      data: [{
        name: "Revenue ($M)",
        labels: years,
        values: deckData.annual.map((r) => +(r.revenue / 1e6).toFixed(1)),
      }],
      options: { barDir: "col", chartColors: [t.colors.STEEL] },
    },
    {
      type: pres.ChartType.line,
      data: [{
        name: "EBITDA %",
        labels: years,
        values: deckData.annual.map((r) => +(r.ebitdaPct * 100).toFixed(1)),
      }],
      options: {
        chartColors: [t.colors.BERRY],
        secondaryValAxis: true, secondaryCatAxis: true,
        lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      },
    },
  ];

  s.addChart(comboData, {
    x: WIDE_POS.x, y: WIDE_POS.y, w: WIDE_POS.w, h: WIDE_POS.h,
    showLegend: true, legendPos: "b",
    legendFontSize: 10, legendFontFace: t.fonts.SANS, legendColor: t.colors.MUTED,
    catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
    valAxes: [
      {
        showValAxisTitle: true, valAxisTitle: "Revenue ($M)",
        valAxisTitleFontFace: t.fonts.SANS, valAxisTitleFontSize: 10, valAxisTitleColor: t.colors.MUTED,
        valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
        valAxisLabelFormatCode: "$0\"M\"",
        valGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
      },
      {
        showValAxisTitle: true, valAxisTitle: "EBITDA %",
        valAxisTitleFontFace: t.fonts.SANS, valAxisTitleFontSize: 10, valAxisTitleColor: t.colors.MUTED,
        valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
        valAxisLabelFormatCode: "0\"%\"",
        valGridLine: { style: "none" },
        valAxisMinVal: 0, valAxisMaxVal: 40,
      },
    ],
    catAxes: [
      { catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED },
      { catAxisHidden: true },
    ],
    plotArea:  { fill: { color: t.colors.BG } },
    chartArea: { fill: { color: t.colors.BG } },
    barGapWidthPct: 55,
  });

  s.addText(
    "Combo charts are where native shines — secondary axis, mixed bar + line, per-series styling all work in one addChart call.",
    {
      x: MARGIN, y: 6.75, w: SLIDE_W - 2 * MARGIN, h: 0.32,
      fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, color: t.colors.INK,
      valign: "top", margin: 0,
    }
  );
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 12 — Donut, stress-test (2025 vs 2035 comparison)
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "DONUT CHART  ·  STRESS-TEST (COMPARISON)",
    "Practice mix, ten years apart.",
    "Two donuts, same color mapping, different years. Tests whether two instances of native addChart can be visually locked together. Center callouts added as overlays — native donuts have no center text option."
  );

  const mixAt = (idx) => ({
    Strategy:   +(deckData.practiceYear.Strategy[idx]   / 1e6).toFixed(2),
    Operations: +(deckData.practiceYear.Operations[idx] / 1e6).toFixed(2),
    Technology: +(deckData.practiceYear.Technology[idx] / 1e6).toFixed(2),
  });
  const mix2025 = mixAt(0);
  const mix2035x = mixAt(10);
  const total = (m) => +(m.Strategy + m.Operations + m.Technology).toFixed(2);

  // Two chart positions, side by side
  const CHART_H = 3.8;
  const CHART_W = 4.5;
  const Y = 2.4;
  const X1 = 1.0;
  const X2 = SLIDE_W - 1.0 - CHART_W;
  const CAPTIONS = [
    { x: X1, title: "2025",     total: total(mix2025), mix: mix2025  },
    { x: X2, title: "2035",     total: total(mix2035x), mix: mix2035x },
  ];

  CAPTIONS.forEach((c) => {
    s.addChart(
      pres.ChartType.doughnut,
      [{
        name: `Revenue mix ${c.title}`,
        labels: ["Strategy", "Operations", "Technology"],
        values: [c.mix.Strategy, c.mix.Operations, c.mix.Technology],
      }],
      {
        x: c.x, y: Y, w: CHART_W, h: CHART_H,
        chartColors: [t.colors.STEEL, t.colors.LBLUE, t.colors.BERRY],
        showLegend: false,
        showValue: false, showPercent: false, showCategoryName: false,
        holeSize: 65,
        plotArea:  { fill: { color: t.colors.BG } },
        chartArea: { fill: { color: t.colors.BG } },
      }
    );

    // Year label under each donut
    s.addText(c.title, {
      x: c.x, y: Y + CHART_H + 0.05, w: CHART_W, h: 0.4,
      fontFace: t.fonts.DISPLAY, fontSize: 20, color: t.colors.INK,
      align: "center", valign: "top", margin: 0,
    });
    // Total at center of donut (since native donuts can't do center text)
    s.addText(`$${c.total.toFixed(1)}M`, {
      x: c.x, y: Y + CHART_H / 2 - 0.18, w: CHART_W, h: 0.4,
      fontFace: t.fonts.DISPLAY, fontSize: 18, bold: false, color: t.colors.INK,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Middle "→" connector + shift narrative between the two donuts
  const midX = (X1 + CHART_W + X2) / 2;
  s.addText("→", {
    x: midX - 0.4, y: Y + CHART_H / 2 - 0.3, w: 0.8, h: 0.6,
    fontFace: t.fonts.DISPLAY, fontSize: 36, color: t.colors.MUTED,
    align: "center", valign: "middle", margin: 0,
  });

  s.addText(
    "Strategy went from 38% of revenue to 31%; Operations grew to the #1 share; Technology more than tripled in dollars.",
    {
      x: MARGIN, y: 6.75, w: SLIDE_W - 2 * MARGIN, h: 0.32,
      fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, color: t.colors.INK,
      valign: "top", margin: 0,
    }
  );
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 13 — Scatter, stress-test (phased multi-series)
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "SCATTER PLOT  ·  STRESS-TEST (PHASED MULTI-SERIES)",
    "Revenue vs EBITDA $, colored by phase.",
    "Four series (setup, compression, recovery, scale) sharing one x-axis. Tests pptxgenjs scatter's null-padding pattern for multi-series with non-overlapping x values."
  );

  // Phase classification (same as build_showcase)
  const phaseOf = (year) => {
    if (year === 2025) return "setup";
    if (year >= 2026 && year <= 2029) return "compression";
    if (year >= 2030 && year <= 2031) return "recovery";
    return "scale";
  };
  const phases = ["setup", "compression", "recovery", "scale"];
  const phaseColorsMap = {
    setup:       t.colors.MUTED,
    compression: t.colors.BERRY,
    recovery:    t.colors.SLATE,
    scale:       t.colors.STEEL,
  };

  // Data: x = revenue $M, y = ebitda $M per year
  const rows = deckData.annual.map((r) => ({
    year: r.year,
    rev: +(r.revenue / 1e6).toFixed(2),
    eb:  +(r.ebitdaPct * r.revenue / 1e6).toFixed(2),
    phase: phaseOf(r.year),
  }));

  // pptxgenjs scatter shape: [{name:"X-Axis", values:[...]}, {name:"Series", values:[...]}].
  // Multi-series with different x-values use null padding.
  const scatterData = [
    { name: "Revenue ($M)", values: rows.map((r) => r.rev) },
  ];
  phases.forEach((p) => {
    scatterData.push({
      name: p.charAt(0).toUpperCase() + p.slice(1),
      values: rows.map((r) => (r.phase === p ? r.eb : null)),
    });
  });

  s.addChart(pres.ChartType.scatter, scatterData, {
    x: WIDE_POS.x, y: WIDE_POS.y, w: WIDE_POS.w, h: WIDE_POS.h,
    chartColors: phases.map((p) => phaseColorsMap[p]),
    lineSize: 0,
    lineDataSymbol: "circle", lineDataSymbolSize: 11,
    lineDataSymbolLineSize: 1,
    showLegend: true, legendPos: "b",
    legendFontSize: 10, legendFontFace: t.fonts.SANS, legendColor: t.colors.MUTED,
    showValAxisTitle: true, valAxisTitle: "EBITDA ($M)",
    showCatAxisTitle: true, catAxisTitle: "Revenue ($M)",
    valAxisTitleFontFace: t.fonts.SANS, valAxisTitleFontSize: 10, valAxisTitleColor: t.colors.MUTED,
    catAxisTitleFontFace: t.fonts.SANS, catAxisTitleFontSize: 10, catAxisTitleColor: t.colors.MUTED,
    valAxisLabelFormatCode: "$0\"M\"",
    catAxisLabelFormatCode: "$0\"M\"",
    valAxisMinVal: 0, valAxisMaxVal: 18,
    catAxisMinVal: 20, catAxisMaxVal: 60,
    valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 10, valAxisLabelColor: t.colors.MUTED,
    catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
    valGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
    catGridLine: { color: t.colors.RULE, style: "solid", size: 0.5 },
    plotArea:  { fill: { color: t.colors.BG } },
    chartArea: { fill: { color: t.colors.BG } },
  });

  s.addText(
    "Compression years sit in the bottom-left cluster. Scale years walk up the diagonal. Phase coloring surfaces structure the simple scatter hides.",
    {
      x: MARGIN, y: 6.75, w: SLIDE_W - 2 * MARGIN, h: 0.32,
      fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, color: t.colors.INK,
      valign: "top", margin: 0,
    }
  );
}

// ══════════════════════════════════════════════════════════════════════
// SLIDE 14 — Radar, stress-test (filled + many axes)
// ══════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  header(s,
    "RADAR CHART  ·  STRESS-TEST (FILLED OVERLAY)",
    "Same three practices, filled regions.",
    "radarStyle = 'filled' with transparency so all three practices read as overlapping shapes. Tests multi-series overlay legibility and whether native radar respects alpha."
  );

  const radarYears = ["2025", "2027", "2029", "2031", "2033", "2035"];
  const radarIdx   = [0, 2, 4, 6, 8, 10];
  const seriesAt = (name) => ({
    name,
    labels: radarYears,
    values: radarIdx.map((i) => +(deckData.practiceYear[name][i] / 1e6).toFixed(2)),
  });
  const radarSeries = [seriesAt("Strategy"), seriesAt("Operations"), seriesAt("Technology")];

  s.addChart(pres.ChartType.radar, radarSeries, {
    x: SPLIT_LEFT.x, y: SPLIT_LEFT.y, w: SPLIT_LEFT.w, h: SPLIT_LEFT.h,
    chartColors: [t.colors.STEEL, t.colors.LBLUE, t.colors.BERRY],
    chartColorsOpacity: 25,
    radarStyle: "filled",
    lineSize: 1.5,
    showLegend: true, legendPos: "b",
    legendFontSize: 10, legendFontFace: t.fonts.SANS, legendColor: t.colors.MUTED,
    catAxisLabelFontFace: t.fonts.SANS, catAxisLabelFontSize: 10, catAxisLabelColor: t.colors.MUTED,
    valAxisLabelFontFace: t.fonts.SANS, valAxisLabelFontSize: 9, valAxisLabelColor: t.colors.MUTED,
    valAxisMinVal: 0, valAxisMaxVal: 22,
    valAxisLabelFormatCode: "$0\"M\"",
    valGridLine: { color: t.colors.RULE, style: "solid", size: 0.4 },
    plotArea:  { fill: { color: t.colors.BG } },
    chartArea: { fill: { color: t.colors.BG } },
  });

  // Right column — what the filled overlay reveals that the line-only didn't
  const rx = SPLIT_RIGHT.x;
  const ry = SPLIT_RIGHT.y;
  const rw = SPLIT_RIGHT.w;

  addEyebrow(s, t, "WHAT THE FILL REVEALS", rx, ry, rw);
  addInkUnderscore(s, t, rx, ry + 0.3, 1.5);

  s.addText(
    "Filled overlay makes the area under each practice's shape comparable at a glance. Where line-only emphasizes trajectory, filled emphasizes cumulative footprint.",
    {
      x: rx, y: ry + 0.6, w: rw, h: 1.2,
      fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, color: t.colors.MUTED,
      margin: 0, valign: "top", lineSpacing: 15,
    }
  );

  s.addText(
    "Native ceiling shows here: three filled regions on six axes fight for legibility. At four-plus series, filled radar stops scaling.",
    {
      x: rx, y: ry + 2.0, w: rw, h: 1.2,
      fontFace: t.fonts.DISPLAY, fontSize: 12, italic: true, color: t.colors.INK,
      margin: 0, valign: "top", lineSpacing: 15,
    }
  );

  addHairline(s, t, ry + 3.5, { x: rx, w: rw });
  s.addText("native limit: no per-axis scale, no conditional fill.", {
    x: rx, y: ry + 3.6, w: rw, h: 0.3,
    fontFace: t.fonts.SANS, fontSize: 10, color: t.colors.MUTED, italic: true,
    valign: "top", margin: 0,
  });
}

// ── Write ─────────────────────────────────────────────────────────────
const outPath = path.resolve(__dirname, "..", "output", "audit_deck.pptx");
pres.writeFile({ fileName: outPath })
  .then((f) => console.log("Wrote:", f))
  .catch((err) => { console.error(err); process.exit(1); });
