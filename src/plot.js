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

// ── Sankey (shape-based, straight-sided ribbons) ──────────────────────
// Two-column flow: one source → N targets. No node rebalancing; source is
// a single stack, targets are stacked in order. Ribbons are straight-sided
// trapezoids (no Bézier curves — pptxgenjs addShape has no curve path).
//
// opts:
//   slide, theme                required
//   source     { label, value }  single source
//   targets    [{ label, value, color? }]  sums should ≈ source.value
//   position   { x, y, w, h }   bounding region
//   nodeWidth  inches (default 0.6) — width of source/target blocks
//   colors     fallback palette tokens (defaults to variety)
//
function drawSankey(opts = {}) {
  const {
    slide, theme: t, source, targets, position,
    nodeWidth = 0.6,
    colors = ["STEEL", "LBLUE", "BERRY", "SLATE"],
  } = opts;

  if (!slide || !t || !source || !targets || !position) {
    throw new Error("drawSankey: slide, theme, source, targets, position are required");
  }

  const totalTarget = targets.reduce((sum, tg) => sum + tg.value, 0);
  const total = Math.max(source.value, totalTarget);

  const srcX = position.x;
  const tgtX = position.x + position.w - nodeWidth;

  const srcH = position.h * (source.value / total);
  const srcY = position.y + (position.h - srcH) / 2;

  // Source node
  slide.addShape("rect", {
    x: srcX, y: srcY, w: nodeWidth, h: srcH,
    fill: { color: resolveColor(t, "INK") },
    line: { type: "none" },
  });
  slide.addText([
    { text: source.label,                            options: { color: t.colors.WHITE, bold: true } },
    { text: `\n$${source.value.toFixed(1)}M`,        options: { color: t.colors.WHITE } },
  ], {
    x: srcX - 0.1, y: srcY, w: nodeWidth + 0.2, h: srcH,
    fontFace: t.fonts.SANS, fontSize: 10,
    align: "center", valign: "middle", margin: 0,
  });

  // Target nodes stacked vertically, total height proportional to totalTarget
  const targetsH = position.h * (totalTarget / total);
  let targetY = position.y + (position.h - targetsH) / 2;
  const targetRects = [];
  targets.forEach((tg, i) => {
    const h = targetsH * (tg.value / totalTarget);
    const col = resolveColor(t, tg.color || colors[i % colors.length]);
    slide.addShape("rect", {
      x: tgtX, y: targetY, w: nodeWidth, h,
      fill: { color: col },
      line: { type: "none" },
    });
    slide.addText([
      { text: tg.label,                     options: { color: t.colors.WHITE, bold: true } },
      { text: `\n$${tg.value.toFixed(1)}M`, options: { color: t.colors.WHITE } },
    ], {
      x: tgtX - 0.3, y: targetY, w: nodeWidth + 0.6, h,
      fontFace: t.fonts.SANS, fontSize: 10,
      align: "center", valign: "middle", margin: 0,
    });
    targetRects.push({ y: targetY, h, color: col });
    targetY += h;
  });

  // Ribbons — each target's share of the source as a trapezoid
  // connecting the source slice to the target rect.
  let sliceStartY = srcY;
  targets.forEach((tg, i) => {
    const sliceH = srcH * (tg.value / totalTarget);
    const tgt = targetRects[i];

    // Use a 4-sided freeform polygon via `custGeom` — but pptxgenjs doesn't
    // expose that directly. Fall back to a filled parallelogram
    // approximation: draw the ribbon as two triangles + a rect, or just
    // as a single "trapezoid" preset rotated.
    //
    // Cleanest: two triangles sharing a shared edge, mid-x between src
    // and target. This reads like a curved ribbon when drawn thin.
    //
    // Pragmatic: use a single 4-sided polygon via the pptxgenjs
    // freeform shape — `addShape("pie", ...)` won't work. Use
    // addShape("rect", ...) with rotate? No — rotation of axis-aligned
    // rect still produces a rotated rect, not a trapezoid.
    //
    // Working approach: draw the ribbon as a series of thin horizontal
    // rectangles stepped from source-y to target-y (stair-stepped
    // polygon approximation). Ugly up close; OK at slide scale.
    const steps = 40;
    const stepW = (tgtX - (srcX + nodeWidth)) / steps;
    for (let k = 0; k < steps; k++) {
      const frac = k / steps;
      const fracNext = (k + 1) / steps;
      // Linear interpolation of top and bottom edges from source to target
      const topY    = sliceStartY + (tgt.y - sliceStartY) * frac;
      const botY    = (sliceStartY + sliceH) + ((tgt.y + tgt.h) - (sliceStartY + sliceH)) * frac;
      const topYNext = sliceStartY + (tgt.y - sliceStartY) * fracNext;
      const botYNext = (sliceStartY + sliceH) + ((tgt.y + tgt.h) - (sliceStartY + sliceH)) * fracNext;
      const ribbonX  = srcX + nodeWidth + k * stepW;
      const avgTopY  = (topY + topYNext) / 2;
      const avgBotY  = (botY + botYNext) / 2;
      slide.addShape("rect", {
        x: ribbonX,
        y: avgTopY,
        w: stepW + 0.01,   // slight overlap avoids BG seams between strips
        h: avgBotY - avgTopY,
        fill: { color: tgt.color, transparency: 55 },
        line: { type: "none" },
      });
    }
    sliceStartY += sliceH;
  });
}

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

// ── Gauge / Tachometer (shape-based half-dial) ────────────────────────
// Uses pptxgenjs's "blockArc" preset — a partial ring with start/end
// angles. Angles are in OOXML convention (0 = 3 o'clock, counterclockwise)
// which pptxgenjs converts from degrees.
//
// opts:
//   slide, theme              required
//   position   { x, y, w, h } bounding box (square works best)
//   value       number        current value
//   domain     [min, max]    scale
//   segments   [{ threshold, color }]  — optional colored zones
//   label      string        displayed under the gauge
//   valueFormat (v) => string  how to display the numeric value
//
function drawGauge(opts = {}) {
  const {
    slide, theme: t, position,
    value, domain = [0, 100],
    segments = [
      { threshold: 0.33, color: "BERRY" },
      { threshold: 0.66, color: "GOLD"  },
      { threshold: 1.0,  color: "STEEL" },
    ],
    label, valueFormat = (v) => `${v.toFixed(0)}`,
  } = opts;

  if (!slide || !t || !position) {
    throw new Error("drawGauge: slide, theme, position are required");
  }

  // Half-circle gauge geometry
  // Use blockArc: 180° to 360° (OOXML) = left-to-right half, upper semicircle
  const cx = position.x + position.w / 2;
  const cy = position.y + position.h * 0.78;   // shifted down so half-circle sits centered vertically
  const outerR = Math.min(position.w, position.h * 1.55) / 2;
  const thickness = outerR * 0.22;
  const innerR = outerR - thickness;

  // Outer bounding rect for the blockArc (full circle size)
  const arcX = cx - outerR;
  const arcY = cy - outerR;
  const arcW = outerR * 2;
  const arcH = outerR * 2;

  // Draw colored segment arcs.
  // blockArc angle convention in pptxgenjs follows OOXML: start/end in degrees,
  // 0° at 3 o'clock going counterclockwise. For a top half-circle from left
  // (9 o'clock, 180°) to right (3 o'clock, 0°), start = 180, end = 0 going
  // CCW through 270 (top).
  // We split the 180° arc into segment shares.

  let startDeg = 180;
  segments.forEach((seg, i) => {
    const prev = i === 0 ? 0 : segments[i - 1].threshold;
    const frac = seg.threshold - prev;
    const sweep = 180 * frac;
    // CCW: end angle is startDeg + sweep, but we go through the top half
    // which means decreasing in OOXML space → end = startDeg + sweep as
    // degrees CCW from 3 o'clock.
    const endDeg = (startDeg + sweep) % 360;
    slide.addShape("blockArc", {
      x: arcX, y: arcY, w: arcW, h: arcH,
      fill: { color: resolveColor(t, seg.color) },
      line: { type: "none" },
      // blockArc adjustment handles (adj1, adj2) control start/end angles.
      // pptxgenjs exposes these through `rectRadius` in some versions, but
      // for broad compatibility we accept that the preset may render the
      // full half-arc if adj isn't settable. In that case we'd need an
      // alternate path (sequential sector overlays) — flagged as a known
      // limit.
      rotate: startDeg - 180,   // rotate the whole arc so our segment starts at 9 o'clock + startDeg
    });
    startDeg += sweep;
  });

  // Needle: rotated thin rectangle pointing to value
  const norm = Math.max(0, Math.min(1, (value - domain[0]) / (domain[1] - domain[0])));
  const needleAngle = 180 - norm * 180;   // degrees from right = 0, going up and over
  const needleLen = outerR * 0.88;
  // Rotate a vertical rectangle by -needleAngle + 90 to align with angle.
  slide.addShape("rect", {
    x: cx - 0.035, y: cy - needleLen, w: 0.07, h: needleLen,
    fill: { color: t.colors.INK },
    line: { type: "none" },
    rotate: needleAngle - 90,
  });
  // Hub
  slide.addShape("ellipse", {
    x: cx - 0.12, y: cy - 0.12, w: 0.24, h: 0.24,
    fill: { color: t.colors.INK },
    line: { type: "none" },
  });

  // Value text below the dial
  slide.addText(valueFormat(value), {
    x: cx - 1.0, y: cy + 0.1, w: 2.0, h: 0.55,
    fontFace: t.fonts.DISPLAY, fontSize: 28, color: t.colors.INK,
    align: "center", valign: "middle", margin: 0,
  });
  if (label) {
    slide.addText(label, {
      x: cx - 1.5, y: cy + 0.65, w: 3.0, h: 0.3,
      fontFace: t.fonts.SANS, fontSize: 10, bold: true,
      color: t.colors.MUTED, charSpacing: 2,
      align: "center", valign: "top", margin: 0,
    });
  }
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
  drawSankey,
  drawHeatmap,
  drawGauge,
  areaScale,
  resolveColor,
  formatTick,
};
