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

// ── Area-proportional sizing helper (Cleveland-correct bubble sizing) ──
// value in [domain[0], domain[1]] → diameter in [range[0], range[1]], area-scaled.
function areaScale({ value, domain, range }) {
  const [vMin, vMax] = domain;
  const [dMin, dMax] = range;
  const norm = (value - vMin) / (vMax - vMin);
  const area = Math.pow(dMin, 2) + norm * (Math.pow(dMax, 2) - Math.pow(dMin, 2));
  return Math.sqrt(area);
}

module.exports = { PlotContext, areaScale, resolveColor, formatTick };
