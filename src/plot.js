// PlotContext — shape-based positioning engine for pptxgenjs slides.
// Every data-driven position on a chart comes from xToSlide / yToSlide.
// All higher-level primitives (bubble, path, axes, quadrants) are built on them.
//
// Coordinate model:
//   outer box = { x, y, w, h }       — the chart's allocated slide region
//   padding   = { left, right, top, bottom }  — space inside outer box for axis labels/titles
//   plot box  = outer box minus padding       — where data marks are rendered
//
// Data-space → slide-space is linear. Y is inverted (data-high = slide-top).

const { saperia } = require("./theme");

function resolveColor(theme, nameOrHex) {
  if (!nameOrHex) return nameOrHex;
  // Six-digit hex without `#` stays as-is; name lookups happen through theme.colors.
  if (theme.colors[nameOrHex]) return theme.colors[nameOrHex];
  if (theme.quadrantTints[nameOrHex]) return theme.quadrantTints[nameOrHex];
  return nameOrHex;
}

function formatTick(value, format) {
  if (typeof format === "function") return format(value);
  if (format === "pct") return `${Math.round(value * 100)}%`;
  if (format === "pct1") return `${(value * 100).toFixed(1)}%`;
  if (format === "dollar") return `$${value}`;
  if (format === "dollarM") return `$${value}M`;
  if (format === "int") return `${Math.round(value)}`;
  return String(value);
}

class PlotContext {
  constructor(slide, opts = {}) {
    this.slide = slide;
    this.theme = opts.theme || saperia;

    const { xRange, yRange, position, padding = {} } = opts;
    if (!xRange || !yRange) throw new Error("PlotContext: xRange and yRange are required");
    if (!position) throw new Error("PlotContext: position is required");

    this.xRange = xRange;
    this.yRange = yRange;

    this.outer = {
      x: position.x,
      y: position.y,
      w: position.w ?? position.size,
      h: position.h ?? position.size,
    };

    this.padding = {
      left:   padding.left   ?? 0.45,
      right:  padding.right  ?? 0.1,
      top:    padding.top    ?? 0.1,
      bottom: padding.bottom ?? 0.4,
    };
  }

  // ── Plot rect (inner area where data lives) ──────────────────────────
  plotLeft()   { return this.outer.x + this.padding.left; }
  plotRight()  { return this.outer.x + this.outer.w - this.padding.right; }
  plotTop()    { return this.outer.y + this.padding.top; }
  plotBottom() { return this.outer.y + this.outer.h - this.padding.bottom; }
  plotRect() {
    return {
      x: this.plotLeft(), y: this.plotTop(),
      w: this.plotRight() - this.plotLeft(),
      h: this.plotBottom() - this.plotTop(),
    };
  }

  // ── Core mapping: data-space → slide-inches ──────────────────────────
  xToSlide(v) {
    const [min, max] = this.xRange;
    const pL = this.plotLeft(), pR = this.plotRight();
    return pL + ((v - min) / (max - min)) * (pR - pL);
  }
  yToSlide(v) {
    const [min, max] = this.yRange;
    const pT = this.plotTop(), pB = this.plotBottom();
    return pB - ((v - min) / (max - min)) * (pB - pT);
  }

  // ── Frame (outer rect around plot area) ──────────────────────────────
  frame(opts = {}) {
    const { color = "RULE", width = 0.5 } = opts;
    const r = this.plotRect();
    this.slide.addShape("rect", {
      x: r.x, y: r.y, w: r.w, h: r.h,
      fill: { type: "none" },
      line: { color: resolveColor(this.theme, color), width },
    });
  }

  // ── Axes (tick labels + axis titles) ─────────────────────────────────
  // Signature: axes({ x: { ticks, format, title }, y: { ticks, format, title } })
  axes(config = {}) {
    const { x: xCfg, y: yCfg } = config;
    const t = this.theme;

    if (xCfg) {
      const ticks = xCfg.ticks || [];
      const pB = this.plotBottom();
      ticks.forEach((v) => {
        const sx = this.xToSlide(v);
        this.slide.addText(formatTick(v, xCfg.format), {
          x: sx - 0.2, y: pB + 0.04, w: 0.4, h: 0.18,
          fontFace: t.fonts.SANS, fontSize: 8, color: t.colors.MUTED,
          align: "center", valign: "top", margin: 0,
        });
      });
      if (xCfg.title) {
        const r = this.plotRect();
        this.slide.addText(xCfg.title, {
          x: r.x, y: pB + 0.26, w: r.w, h: 0.22,
          fontFace: t.fonts.SANS, fontSize: 9, bold: true, color: t.colors.MUTED,
          charSpacing: 2, align: "center", valign: "top", margin: 0,
        });
      }
    }

    if (yCfg) {
      const ticks = yCfg.ticks || [];
      ticks.forEach((v) => {
        const sy = this.yToSlide(v);
        this.slide.addText(formatTick(v, yCfg.format), {
          x: this.outer.x, y: sy - 0.09,
          w: this.padding.left - 0.08, h: 0.18,
          fontFace: t.fonts.SANS, fontSize: 8, color: t.colors.MUTED,
          align: "right", valign: "top", margin: 0,
        });
      });
      if (yCfg.title) {
        const r = this.plotRect();
        this.slide.addText(yCfg.title, {
          x: this.outer.x - 0.5, y: r.y + r.h / 2, w: 1.0, h: 0.22,
          fontFace: t.fonts.SANS, fontSize: 9, bold: true, color: t.colors.MUTED,
          charSpacing: 2, align: "center", valign: "middle", margin: 0,
          rotate: -90,
        });
      }
    }
  }

  // ── Quadrants (tints + split lines + corner labels) ──────────────────
  // config = {
  //   split: [xVal, yVal],
  //   labels: [
  //     { corner: "tl"|"tr"|"bl"|"br", hdr, sub, tint }
  //   ],
  //   splitLine: { color, width, transparency }     // optional override
  // }
  quadrants(config = {}) {
    const t = this.theme;
    const [xSplit, ySplit] = config.split || [];
    if (xSplit == null || ySplit == null) {
      throw new Error("PlotContext.quadrants: split is required");
    }
    const r = this.plotRect();
    const pL = r.x, pR = r.x + r.w, pT = r.y, pB = r.y + r.h;
    const sx = this.xToSlide(xSplit);
    const sy = this.yToSlide(ySplit);

    // Tints (always placed by corner; caller-supplied tint name)
    const byCorner = {};
    (config.labels || []).forEach((L) => { byCorner[L.corner] = L; });

    const corners = [
      { k: "tl", x: pL, y: pT, w: sx - pL, h: sy - pT, transparency: 75 },
      { k: "tr", x: sx, y: pT, w: pR - sx, h: sy - pT, transparency: 65 },
      { k: "bl", x: pL, y: sy, w: sx - pL, h: pB - sy, transparency: 70 },
      { k: "br", x: sx, y: sy, w: pR - sx, h: pB - sy, transparency: 70 },
    ];
    corners.forEach((c) => {
      const L = byCorner[c.k];
      if (!L || !L.tint) return;
      this.slide.addShape("rect", {
        x: c.x, y: c.y, w: c.w, h: c.h,
        fill: { color: resolveColor(t, L.tint), transparency: L.transparency ?? c.transparency },
        line: { type: "none" },
      });
    });

    // Split lines
    const sl = config.splitLine || {};
    this.slide.addShape("line", {
      x: sx, y: pT, w: 0, h: pB - pT,
      line: {
        color: resolveColor(t, sl.color || "INK"),
        width: sl.width ?? 0.75,
        transparency: sl.transparency ?? 60,
      },
    });
    this.slide.addShape("line", {
      x: pL, y: sy, w: pR - pL, h: 0,
      line: {
        color: resolveColor(t, sl.color || "INK"),
        width: sl.width ?? 0.75,
        transparency: sl.transparency ?? 60,
      },
    });

    // Corner labels
    const lf = {
      fontFace: t.fonts.SANS, fontSize: 10, bold: true, color: t.colors.INK,
      charSpacing: 1.5, margin: 0, valign: "top",
    };
    const df = {
      fontFace: t.fonts.DISPLAY, fontSize: 9, italic: true, color: t.colors.INK,
      margin: 0, valign: "top",
    };
    const cornerOrigin = {
      tl: { hx: pL + 0.08, hy: pT + 0.08, sy: pT + 0.28 },
      tr: { hx: sx + 0.08, hy: pT + 0.08, sy: pT + 0.28 },
      bl: { hx: pL + 0.08, hy: pB - 0.34, sy: pB - 0.18 },
      br: { hx: sx + 0.08, hy: pB - 0.34, sy: pB - 0.18 },
    };
    (config.labels || []).forEach((L) => {
      const o = cornerOrigin[L.corner];
      if (!o) return;
      if (L.hdr) {
        this.slide.addText(L.hdr, { x: o.hx, y: o.hy, w: 1.8, h: 0.22, ...lf });
      }
      if (L.sub) {
        this.slide.addText(L.sub, { x: o.hx, y: o.sy, w: 2.0, h: 0.2, ...df });
      }
    });
  }

  // ── Bubble (ellipse at data coords, sized in slide-inch diameter) ────
  bubble(opts = {}) {
    const t = this.theme;
    const { x, y, size, color = "INK", transparency = 25, lineWidth = 1.25,
            label, labelColor = "WHITE", labelFontSize = 8 } = opts;

    const cx = this.xToSlide(x);
    const cy = this.yToSlide(y);
    const col = resolveColor(t, color);

    this.slide.addShape("ellipse", {
      x: cx - size / 2, y: cy - size / 2, w: size, h: size,
      fill: { color: col, transparency },
      line: { color: col, width: lineWidth },
    });

    if (label != null) {
      this.slide.addText(String(label), {
        x: cx - 0.22, y: cy - 0.09, w: 0.44, h: 0.18,
        fontFace: t.fonts.SANS, fontSize: labelFontSize,
        color: resolveColor(t, labelColor),
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    }
  }

  // ── Path (connected line segments between data points) ───────────────
  // points = [{x, y}, ...]
  path(points, opts = {}) {
    const t = this.theme;
    const { color = "MUTED", width = 0.5, transparency = 55 } = opts;
    const col = resolveColor(t, color);

    for (let i = 0; i < points.length - 1; i++) {
      const a = points[i], b = points[i + 1];
      const ax = this.xToSlide(a.x), ay = this.yToSlide(a.y);
      const bx = this.xToSlide(b.x), by = this.yToSlide(b.y);
      this.slide.addShape("line", {
        x: Math.min(ax, bx), y: Math.min(ay, by),
        w: Math.abs(bx - ax), h: Math.abs(by - ay),
        line: { color: col, width, transparency },
        flipH: ax > bx, flipV: ay > by,
      });
    }
  }

  // ── Escape hatch: positioned shape at data coords ────────────────────
  markShape(shape, opts = {}) {
    const { x, y, w = 0, h = 0, ...rest } = opts;
    const cx = this.xToSlide(x);
    const cy = this.yToSlide(y);
    this.slide.addShape(shape, { x: cx - w / 2, y: cy - h / 2, w, h, ...rest });
  }

  // ── Bar (rect at data-x center, spanning baseline→value on data-y) ───
  // Category charts: pass xRange like [-0.5, n-0.5] and set x to the
  // category index (0, 1, 2, ...). For float bars (waterfall), set
  // `baseline` to the bar's starting value; defaults to yRange[0] (floor).
  bar(opts = {}) {
    const t = this.theme;
    const {
      x, value,
      width = 0.5,                      // bar width in slide inches
      color = "STEEL",
      baseline = this.yRange[0],
      transparency = 0,
      lineWidth = 0,
      label,
      labelColor = "INK",
      labelFontSize = 9,
      labelOffset = 0.08,               // distance above top of bar
    } = opts;

    const col = resolveColor(t, color);
    const cx = this.xToSlide(x);
    const topData    = Math.max(value, baseline);
    const bottomData = Math.min(value, baseline);
    const yTop       = this.yToSlide(topData);
    const yBottom    = this.yToSlide(bottomData);

    this.slide.addShape("rect", {
      x: cx - width / 2,
      y: yTop,
      w: width,
      h: yBottom - yTop,
      fill: { color: col, transparency },
      line: lineWidth > 0
        ? { color: col, width: lineWidth }
        : { type: "none" },
    });

    if (label != null) {
      this.slide.addText(String(label), {
        x: cx - width, y: yTop - 0.32 - labelOffset,
        w: width * 2, h: 0.3,
        fontFace: t.fonts.SANS, fontSize: labelFontSize,
        color: resolveColor(t, labelColor),
        bold: true, align: "center", valign: "bottom", margin: 0,
      });
    }
  }

  // ── Grid (horizontal or vertical lines at data-space tick values) ────
  // grid({ y: [ticks], x: [ticks], color, width, transparency, skipFirst })
  // `skipFirst: true` (default) suppresses the line at the first tick, so
  // the baseline isn't double-drawn over the frame / x-axis.
  grid(opts = {}) {
    const t = this.theme;
    const {
      y: yTicks, x: xTicks,
      color = "RULE", width = 0.4, transparency = 40,
      skipFirst = true,
    } = opts;
    const col = resolveColor(t, color);
    const r = this.plotRect();

    if (yTicks) {
      const ticks = skipFirst ? yTicks.slice(1) : yTicks;
      ticks.forEach((v) => {
        const sy = this.yToSlide(v);
        this.slide.addShape("line", {
          x: r.x, y: sy, w: r.w, h: 0,
          line: { color: col, width, transparency },
        });
      });
    }

    if (xTicks) {
      const ticks = skipFirst ? xTicks.slice(1) : xTicks;
      ticks.forEach((v) => {
        const sx = this.xToSlide(v);
        this.slide.addShape("line", {
          x: sx, y: r.y, w: 0, h: r.h,
          line: { color: col, width, transparency },
        });
      });
    }
  }
}

// ── Funnel (standalone — not a PlotContext method) ────────────────────
// Draws a vertically-stacked trapezoid funnel inside a slide rectangle.
// Not inside PlotContext because funnels have no xRange/yRange — they're
// pure proportional composition, self-contained.
//
// opts:
//   slide        required
//   theme        required (normally saperia)
//   stages       [{label, value, sub?}]   top to bottom
//   position     {x, y, w, h}             the box the funnel sits in
//   color        token name (default "STEEL")
//   narrowTo     fraction 0-1 of widest — width of the narrowest stage
//                relative to the widest (default 0.25)
//   gap          slide-inches between stages (default 0.06)
//   showLabels   bool, default true — inline stage labels + values
//
function drawFunnel(opts = {}) {
  const {
    slide, theme: t, stages, position,
    color = "STEEL", narrowTo = 0.25, gap = 0.06,
    showLabels = true,
  } = opts;

  if (!slide || !t || !stages || !position) {
    throw new Error("drawFunnel: slide, theme, stages, position are required");
  }

  const col = resolveColor(t, color);
  const n = stages.length;
  const stageH = (position.h - gap * (n - 1)) / n;
  const maxV = Math.max(...stages.map((s) => s.value));
  const minV = Math.min(...stages.map((s) => s.value));

  // Map a stage's value to a normalized width in [narrowTo, 1].
  const valueToWidth = (v) =>
    narrowTo + ((v - minV) / (maxV - minV || 1)) * (1 - narrowTo);

  const cx = position.x + position.w / 2;

  stages.forEach((stg, i) => {
    const topV    = stg.value;
    const botV    = i < n - 1 ? stages[i + 1].value : stg.value * narrowTo;
    const topW    = position.w * valueToWidth(topV);
    const botW    = i < n - 1
      ? position.w * valueToWidth(botV)
      : topW * 0.55;   // final stage narrows further to close the funnel
    const y       = position.y + i * (stageH + gap);

    // Trapezoid via custGeom — pptxgenjs supports custom geometry.
    // Fallback: build with four line segments + a poly-shape via rect +
    // two triangle corners. Cleanest path: the "trapezoid" prstGeom.
    slide.addShape("trapezoid", {
      x: cx - topW / 2,
      y,
      w: topW,
      h: stageH,
      fill: { color: col, transparency: 18 + i * 8 },   // fade down
      line: { color: col, width: 0.5 },
      flipV: true,   // prstGeom trapezoid points up by default; flip to widen-at-top
    });

    if (showLabels) {
      // Stage label + value, centered inside the trapezoid
      slide.addText(
        [
          { text: stg.label,              options: { color: t.colors.WHITE, bold: true } },
          { text: `  ·  ${stg.value}`,    options: { color: t.colors.WHITE } },
        ],
        {
          x: cx - topW / 2, y: y + stageH / 2 - 0.15,
          w: topW, h: 0.3,
          fontFace: t.fonts.SANS, fontSize: 11,
          align: "center", valign: "middle", margin: 0,
        }
      );
      if (stg.sub) {
        slide.addText(stg.sub, {
          x: cx - topW / 2, y: y + stageH / 2 + 0.08,
          w: topW, h: 0.28,
          fontFace: t.fonts.DISPLAY, fontSize: 10, italic: true,
          color: t.colors.WHITE, align: "center", valign: "top", margin: 0,
        });
      }
    }
  });
}

// ── Treemap (standalone) ──────────────────────────────────────────────
// Slice-and-dice layout. For each recursion, split along the longer
// dimension of the remaining rectangle; largest-value item gets the
// share proportional to its value. Not squarified, but deterministic,
// compact, and fine at 3–20 items.
//
// opts:
//   slide        required
//   theme        required
//   items        [{label, value, color?, sub?}]  — will be sorted desc
//   position     {x, y, w, h}
//   gap          slide-inches between rects (default 0.04)
//   colors       array of palette tokens if items don't specify color
//   labelMin     skip label if rect w or h smaller than this (default 0.6)
//
function drawTreemap(opts = {}) {
  const {
    slide, theme: t, items, position,
    gap = 0.04,
    colors = ["STEEL", "LBLUE", "BERRY", "SLATE", "MUTED", "GOLD"],
    labelMin = 0.6,
  } = opts;

  if (!slide || !t || !items || !position) {
    throw new Error("drawTreemap: slide, theme, items, position are required");
  }

  const sorted = [...items]
    .map((it, i) => ({ ...it, _idx: i }))
    .sort((a, b) => b.value - a.value);

  const total = sorted.reduce((s, it) => s + it.value, 0);

  // Recursive slice-and-dice.
  function layout(remaining, rect) {
    if (remaining.length === 0) return [];
    if (remaining.length === 1) {
      return [{ item: remaining[0], rect }];
    }
    const first = remaining[0];
    const rest  = remaining.slice(1);
    const sumHere = remaining.reduce((s, it) => s + it.value, 0);
    const frac = first.value / sumHere;

    let firstRect, restRect;
    if (rect.w >= rect.h) {
      // Split horizontally: first takes left portion
      const firstW = rect.w * frac;
      firstRect = { x: rect.x, y: rect.y, w: firstW - gap, h: rect.h };
      restRect  = { x: rect.x + firstW, y: rect.y, w: rect.w - firstW, h: rect.h };
    } else {
      // Split vertically: first takes top portion
      const firstH = rect.h * frac;
      firstRect = { x: rect.x, y: rect.y, w: rect.w, h: firstH - gap };
      restRect  = { x: rect.x, y: rect.y + firstH, w: rect.w, h: rect.h - firstH };
    }
    return [{ item: first, rect: firstRect }, ...layout(rest, restRect)];
  }

  const placed = layout(sorted, { ...position });

  placed.forEach((p, i) => {
    const col = resolveColor(
      t,
      p.item.color || colors[p.item._idx % colors.length]
    );
    slide.addShape("rect", {
      x: p.rect.x, y: p.rect.y, w: p.rect.w, h: p.rect.h,
      fill: { color: col },
      line: { color: t.colors.BG, width: 1 },
    });
    if (p.rect.w >= labelMin && p.rect.h >= labelMin && p.item.label) {
      slide.addText(p.item.label, {
        x: p.rect.x + 0.06, y: p.rect.y + 0.06,
        w: p.rect.w - 0.12, h: 0.3,
        fontFace: t.fonts.SANS, fontSize: 10, bold: true, color: t.colors.WHITE,
        charSpacing: 1, valign: "top", margin: 0,
      });
      if (p.item.sub) {
        slide.addText(p.item.sub, {
          x: p.rect.x + 0.06, y: p.rect.y + 0.36,
          w: p.rect.w - 0.12, h: 0.3,
          fontFace: t.fonts.DISPLAY, fontSize: 11, italic: true, color: t.colors.WHITE,
          valign: "top", margin: 0,
        });
      }
      if (p.item.value != null && p.rect.h >= 0.8) {
        slide.addText(`$${(p.item.value).toFixed(1)}M`, {
          x: p.rect.x + 0.06, y: p.rect.y + p.rect.h - 0.36,
          w: p.rect.w - 0.12, h: 0.3,
          fontFace: t.fonts.DISPLAY, fontSize: 14, color: t.colors.WHITE,
          valign: "bottom", margin: 0,
        });
      }
    }
  });
}

// ── Sankey: removed in round 5. ───────────────────────────────────────
// Shape-based straight-ribbon approximations render as clunky at slide
// scale. For Sankey, use matplotlib PNG + addImage (see handoff's
// make_sankey.py pattern).

// ── Heatmap (shape-based grid with interpolated cell fills) ───────────
// opts:
//   slide, theme                required
//   data       2D array [rows][cols]
//   rowLabels  array of row labels (length = rows)
//   colLabels  array of col labels (length = cols)
//   position   { x, y, w, h }  — includes space for labels
//   colorFrom  token for low end (default "BG_RAISED")
//   colorTo    token for high end (default "STEEL")
//   domain     [min, max] — defaults to data min/max
//   labelCells bool, default true — print value in each cell
//
function drawHeatmap(opts = {}) {
  const {
    slide, theme: t, data, rowLabels, colLabels, position,
    colorFrom = "BG_RAISED", colorTo = "STEEL",
    domain,
    labelCells = true,
    valueFormat = (v) => v.toFixed(1),
  } = opts;

  if (!slide || !t || !data || !position) {
    throw new Error("drawHeatmap: slide, theme, data, position are required");
  }

  const rows = data.length;
  const cols = data[0].length;

  const labelW = 1.1;    // left gutter for row labels
  const labelH = 0.28;   // top strip for col labels
  const cellX0 = position.x + labelW;
  const cellY0 = position.y + labelH;
  const cellW  = (position.w - labelW) / cols;
  const cellH  = (position.h - labelH) / rows;

  // Flatten values to compute domain
  const allVals = data.flat();
  const [vMin, vMax] = domain || [Math.min(...allVals), Math.max(...allVals)];

  const fromRGB = hexToRgb(resolveColor(t, colorFrom));
  const toRGB   = hexToRgb(resolveColor(t, colorTo));

  // Column labels
  if (colLabels) {
    colLabels.forEach((lbl, c) => {
      slide.addText(String(lbl), {
        x: cellX0 + c * cellW, y: position.y, w: cellW, h: labelH,
        fontFace: t.fonts.SANS, fontSize: 9, color: t.colors.MUTED,
        align: "center", valign: "bottom", margin: 0,
      });
    });
  }

  // Row labels + cells
  data.forEach((row, r) => {
    if (rowLabels) {
      slide.addText(String(rowLabels[r]), {
        x: position.x, y: cellY0 + r * cellH, w: labelW - 0.1, h: cellH,
        fontFace: t.fonts.SANS, fontSize: 10, bold: true, color: t.colors.INK,
        align: "right", valign: "middle", margin: 0,
      });
    }
    row.forEach((v, c) => {
      const norm = vMax === vMin ? 0 : (v - vMin) / (vMax - vMin);
      const rgb = [
        Math.round(fromRGB[0] + (toRGB[0] - fromRGB[0]) * norm),
        Math.round(fromRGB[1] + (toRGB[1] - fromRGB[1]) * norm),
        Math.round(fromRGB[2] + (toRGB[2] - fromRGB[2]) * norm),
      ];
      const cellColor = rgbToHex(rgb);
      slide.addShape("rect", {
        x: cellX0 + c * cellW, y: cellY0 + r * cellH, w: cellW, h: cellH,
        fill: { color: cellColor },
        line: { color: t.colors.BG, width: 1 },
      });
      if (labelCells) {
        const textColor = norm > 0.55 ? t.colors.WHITE : t.colors.INK;
        slide.addText(valueFormat(v), {
          x: cellX0 + c * cellW, y: cellY0 + r * cellH, w: cellW, h: cellH,
          fontFace: t.fonts.SANS, fontSize: 9,
          color: textColor,
          align: "center", valign: "middle", margin: 0,
        });
      }
    });
  });
}

function hexToRgb(hex) {
  const h = hex.length === 3
    ? hex.split("").map((c) => c + c).join("")
    : hex;
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}
function rgbToHex([r, g, b]) {
  const h = (n) => n.toString(16).padStart(2, "0").toUpperCase();
  return `${h(r)}${h(g)}${h(b)}`;
}

// ── Gauge (hybrid: matplotlib arc PNG + native overlays) ──────────────
// Minimum-outsourcing approach: the matplotlib PNG contains ONLY the
// three colored arc zones (the curvy part pptxgenjs can't do cleanly).
// Everything else — tick marks, tick labels, target triangle, needle,
// hub, all text — is rendered natively so it remains editable in
// PowerPoint and each gauge can be positioned independently.
//
// opts:
//   slide, theme                 required
//   position { x, y, w, h }      the whole gauge card (title + arc + value stack)
//   title, subtitle              rendered natively above the arc
//   value, target                data (numbers)
//   domain    [lo, hi]           arc scale
//   ticks     [v, v, ...]        value positions to label on the arc
//   higherIsBetter  true | false | null
//                                — true: red-to-green left-to-right (default)
//                                — false: green-to-red (flipH the arc PNG)
//                                — null:  same as true, but no good/bad framing
//   valueFormat (v) => string    how the center value is displayed
//   targetFormat (v) => string   how the target line reads
//   arcImagePath  string         absolute path to gauge_arc.png
//
function drawGauge(opts = {}) {
  const {
    slide, theme: t,
    position,
    title, subtitle,
    value, target, domain,
    ticks = [],
    higherIsBetter = true,
    valueFormat = (v) => `${v}`,
    targetFormat = null,
    arcImagePath,
  } = opts;

  if (!slide || !t || !position || !arcImagePath) {
    throw new Error("drawGauge: slide, theme, position, arcImagePath are required");
  }
  const [lo, hi] = domain;

  const { x, y, w, h } = position;

  // Header: title + subtitle ------------------------------------------
  const headerTop   = y;
  const titleH      = 0.28;
  const subtitleH   = 0.22;
  slide.addText(title || "", {
    x, y: headerTop, w, h: titleH,
    fontFace: t.fonts.SANS, fontSize: 12, bold: true, color: t.colors.INK,
    align: "center", valign: "middle", margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x, y: headerTop + titleH, w, h: subtitleH,
      fontFace: t.fonts.DISPLAY, fontSize: 10, italic: true, color: t.colors.MUTED,
      align: "center", valign: "middle", margin: 0,
    });
  }

  // Arc region ---------------------------------------------------------
  // The PNG is a 2:1 aspect semicircle (a wide half-circle).
  // Fit it inside the remaining card width, reserving space below for
  // the value stack (value + target + delta).
  const headerBottomY = headerTop + titleH + (subtitle ? subtitleH : 0);
  const valueStackH   = 1.05;   // value (22pt) + target + delta + gap
  const arcSlotH      = h - (headerBottomY - y) - valueStackH;
  const maxArcH = Math.max(0.6, arcSlotH);
  // Arc PNG width constrained by both slot width and 2× slot height
  const arcW = Math.min(w * 0.88, maxArcH * 2);
  const arcH = arcW / 2;
  const arcX = x + (w - arcW) / 2;
  const arcY = headerBottomY + 0.05;

  slide.addImage({
    path: arcImagePath,
    x: arcX, y: arcY, w: arcW, h: arcH,
    flipH: higherIsBetter === false,
  });

  // Arc geometry for overlays
  const cx      = arcX + arcW / 2;
  const cy      = arcY + arcH;
  const rOuter  = arcW / 2;
  const trackW  = rOuter * 0.22;
  const rInner  = rOuter - trackW;
  const rMid    = (rOuter + rInner) / 2;

  const angleAt = (v) => 180 - ((v - lo) / (hi - lo)) * 180;   // degrees, CCW from +x

  // Tick marks + labels ----------------------------------------------
  const tickLen = rOuter * 0.08;
  ticks.forEach((tv) => {
    const thetaDeg = angleAt(tv);
    const thetaRad = thetaDeg * Math.PI / 180;
    const x1 = cx + Math.cos(thetaRad) * rOuter;
    const y1 = cy - Math.sin(thetaRad) * rOuter;
    const x2 = cx + Math.cos(thetaRad) * (rOuter + tickLen);
    const y2 = cy - Math.sin(thetaRad) * (rOuter + tickLen);
    slide.addShape("line", {
      x: Math.min(x1, x2), y: Math.min(y1, y2),
      w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      line: { color: t.colors.MUTED, width: 0.5 },
      flipH: x1 > x2, flipV: y1 > y2,
    });
    // Label further out
    const labelR = rOuter + tickLen + 0.16;
    const lx = cx + Math.cos(thetaRad) * labelR;
    const ly = cy - Math.sin(thetaRad) * labelR;
    slide.addText(String(tv), {
      x: lx - 0.22, y: ly - 0.11, w: 0.44, h: 0.22,
      fontFace: t.fonts.SANS, fontSize: 8, color: t.colors.MUTED,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Min / max labels at the ends of the arc
  slide.addText(String(lo), {
    x: cx - rOuter - 0.26, y: cy - 0.08, w: 0.25, h: 0.22,
    fontFace: t.fonts.SANS, fontSize: 8, color: t.colors.MUTED,
    align: "right", valign: "middle", margin: 0,
  });
  slide.addText(String(hi), {
    x: cx + rOuter + 0.01, y: cy - 0.08, w: 0.25, h: 0.22,
    fontFace: t.fonts.SANS, fontSize: 8, color: t.colors.MUTED,
    align: "left", valign: "middle", margin: 0,
  });

  // Target triangle ---------------------------------------------------
  // Placed past the tick labels, apex pointing toward the gauge center.
  // pptxgenjs triangle default points UP; rotate so apex points toward origin.
  // Rotation formula (derived): (90 + tFrac * 180) mod 360, clockwise.
  const tFrac = Math.max(0, Math.min(1, (target - lo) / (hi - lo)));
  const tDeg  = angleAt(target);
  const tRad  = tDeg * Math.PI / 180;
  const triR  = rOuter + tickLen + 0.38;
  const triX  = cx + Math.cos(tRad) * triR;
  const triY  = cy - Math.sin(tRad) * triR;
  const triSize = 0.14;
  const triRot  = (90 + tFrac * 180) % 360;
  slide.addShape("triangle", {
    x: triX - triSize / 2, y: triY - triSize / 2,
    w: triSize, h: triSize,
    fill: { color: t.colors.STEEL },
    line: { type: "none" },
    rotate: triRot,
  });

  // Needle ------------------------------------------------------------
  const vFrac = Math.max(0, Math.min(1, (value - lo) / (hi - lo)));
  const vDeg  = angleAt(value);
  const vRad  = vDeg * Math.PI / 180;
  const tipR  = rMid + 0.04;
  const tipX  = cx + Math.cos(vRad) * tipR;
  const tipY  = cy - Math.sin(vRad) * tipR;
  slide.addShape("line", {
    x: Math.min(cx, tipX), y: Math.min(cy, tipY),
    w: Math.abs(tipX - cx), h: Math.abs(tipY - cy),
    line: { color: t.colors.BERRY, width: 2.2 },
    flipH: cx > tipX, flipV: cy > tipY,
  });

  // Hub ---------------------------------------------------------------
  const hubR = 0.11;
  slide.addShape("ellipse", {
    x: cx - hubR, y: cy - hubR, w: hubR * 2, h: hubR * 2,
    fill: { color: t.colors.INK },
    line: { type: "none" },
  });

  // Value + target + delta stack --------------------------------------
  const stackY = cy + 0.14;
  slide.addText(valueFormat(value), {
    x, y: stackY, w, h: 0.5,
    fontFace: t.fonts.DISPLAY, fontSize: 22, color: t.colors.INK,
    align: "center", valign: "top", margin: 0,
  });

  const tText = targetFormat
    ? targetFormat(target)
    : (higherIsBetter === false ? `Target: <${target}%` : `Target: ${target}`);
  slide.addText(tText, {
    x, y: stackY + 0.48, w, h: 0.24,
    fontFace: t.fonts.DISPLAY, fontSize: 9, italic: true, color: t.colors.INK,
    align: "center", valign: "top", margin: 0,
  });

  // Delta line — status-colored
  const delta = value - target;
  const statusGood = higherIsBetter === true ? (value >= target) :
                     higherIsBetter === false ? (value <= target) :
                     Math.abs(delta) / Math.max(1, target) < 0.25;
  const deltaColor = statusGood ? t.colors.STEEL : t.colors.BERRY;
  let deltaStr;
  if (higherIsBetter === false) {
    const below = target - value;
    const sign = below >= 0 ? "-" : "+";
    deltaStr = `${sign}${Math.abs(below).toFixed(1)} below target`;
  } else {
    const sign = delta >= 0 ? "+" : "";
    const side = delta >= 0 ? "above" : "below";
    deltaStr = `${sign}${delta.toFixed(1)} ${side} target`;
  }
  slide.addText(deltaStr, {
    x, y: stackY + 0.75, w, h: 0.24,
    fontFace: t.fonts.SANS, fontSize: 9, bold: true, color: deltaColor,
    align: "center", valign: "top", margin: 0,
  });
}

// ── Area-proportional sizing helper (Cleveland-correct bubble sizing) ──
// value in [domain[0], domain[1]] → diameter in [range[0], range[1]], area-scaled.
function areaScale({ value, domain, range }) {
  const [vMin, vMax] = domain;
  const [dMin, dMax] = range;
  const norm = (value - vMin) / (vMax - vMin);
  const area = Math.pow(dMin, 2) + norm * (Math.pow(dMax, 2) - Math.pow(dMin, 2));
  return Math.sqrt(area);
}

module.exports = {
  PlotContext,
  drawFunnel,
  drawTreemap,
  drawHeatmap,
  drawGauge,
  areaScale,
  resolveColor,
  formatTick,
};
