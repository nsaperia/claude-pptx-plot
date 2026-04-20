# pptxgenjs native chart capabilities

Reference for what you can and can't do with `slide.addChart()` in pptxgenjs **3.12.x**. Organized by category with short notes on edge cases and known limits. Companion to `output/audit_deck.pptx` — the deck shows these in action; this doc is the index.

**Hex color rule:** every `color:` in pptxgenjs options takes a 6-digit hex **without** the `#` prefix. Including it renders shapes wrong.

---

## Chart types (via `pres.ChartType`)

| Type | Native? | Notes |
|---|---|---|
| `bar` | ✓ | Default is horizontal. `barDir: "col"` for vertical columns. |
| `line` | ✓ | Markers off by default; turn on via `lineDataSymbol`. |
| `scatter` | ✓ | Non-bubble. Data shape `[{name:"X", values:[...]}, {name:"Series", values:[...]}]`. Multi-series uses null padding for non-overlapping x. |
| `bubble` | ✓ | Data adds `sizes:[...]`. **Cannot override bubble scale** — pptxgenjs doesn't expose `<c:bubbleScale>`; sizes are ratios only. |
| `pie` | ✓ | — |
| `doughnut` | ✓ | `holeSize: 10-90`. No center-text option — overlay with `addText`. |
| `radar` | ✓ | `radarStyle: "standard" \| "marker" \| "filled"`. |
| `area` | ✓ | Stacking via `barGrouping`. |
| Combo (bar + line, etc.) | ✓ | Pass an array of typed blocks to `addChart(dataArr, options)` instead of a single type. Each block has its own type/data/options. |
| Waterfall | ✗ | No native type. Workaround: stacked column with invisible base series — loses per-bar color. |
| Sankey / Treemap / Sunburst | ✗ | Not supported. Use shapes (`claude-pptx-plot`) or matplotlib PNG insert. |
| Gantt | ✗ | Shape-based only. |
| Funnel | ✗ | Shape-based only. |

---

## Common options (all chart types)

### Positioning / sizing

- `x`, `y`, `w`, `h` — slide inches (with `LAYOUT_WIDE`: 13.33 × 7.5).

### Fill

- `plotArea.fill`  — the inner plotting region
- `chartArea.fill` — the full chart bounding box

Both take `{ color, transparency }`. Match slide BG to blend seamlessly.

### Legend

- `showLegend: boolean`
- `legendPos: "t" | "b" | "l" | "r" | "tr"` — top, bottom, left, right, top-right
- `legendFontFace`, `legendFontSize`, `legendColor`

### Grid

- `valGridLine: { color, style, size }` — `style: "solid" | "dash" | "dashDot" | "none"`
- `catGridLine: { ... }` — same fields

### Axes (single axis case)

- `catAxisHidden`, `valAxisHidden`
- `catAxisLineShow`, `valAxisLineShow`
- `catAxisLabelFontFace`, `catAxisLabelFontSize`, `catAxisLabelColor`
- `valAxisLabelFontFace`, `valAxisLabelFontSize`, `valAxisLabelColor`
- `valAxisLabelFormatCode` — Excel format strings: `"0"`, `"0.0%"`, `"$0"`, `"$0\"M\""`, `"0\"%\""`
- `catAxisLabelFormatCode` — same
- `valAxisMinVal`, `valAxisMaxVal`
- `showValAxisTitle`, `valAxisTitle`, `valAxisTitleFontFace`, `valAxisTitleFontSize`, `valAxisTitleColor`
- `showCatAxisTitle`, `catAxisTitle`, `catAxisTitleFontFace`, `catAxisTitleFontSize`, `catAxisTitleColor`
- `valAxisLogScaleBase` — pass a base (e.g. `10`) for log scale

### Axes (combo / dual-axis case)

Instead of singular props, pass arrays:

- `valAxes: [{...primary}, {...secondary}]`
- `catAxes: [{...primary}, {...secondary}]`

On a combo chart's secondary block, set `secondaryValAxis: true` and `secondaryCatAxis: true` in that block's options.

### Data labels (where shown at all)

- `showValue`, `showPercent`, `showCategoryName`, `showLegendKey` — booleans, often mutually combinable
- `dataLabelFontFace`, `dataLabelFontSize`, `dataLabelColor`
- `dataLabelPosition`: varies by type — `"ctr" | "inEnd" | "outEnd" | "bestFit" | "b" | "t" | "l" | "r"`
- `dataLabelFormatCode` — same Excel format strings

**Limit:** no per-point label styling. Entire series gets the same font/color/format.

---

## Series colors

- `chartColors: ["STEEL", "BERRY", ...]` (hex strings, no `#`). One entry per series.
- `chartColorsOpacity: 0-100` — applied to all series uniformly.

**Limit:** no per-point color. For a single bar-chart series, every bar is the same color. Workarounds:
- Split into N single-point series (legend gets ugly)
- Escape to pure shapes via `claude-pptx-plot`

---

## Bar / Column specific

- `barDir: "col" | "bar"` — vertical vs horizontal
- `barGrouping: "clustered" | "stacked" | "percentStacked" | "standard"`
- `barGapWidthPct: 0-500` — gap between groups (%)
- `barOverlapPct: -100 to 100` — overlap within group

---

## Line / Scatter specific

- `lineSize: number` — stroke width in pt. Set to `0` to hide the line (scatter dots only).
- `lineDataSymbol: "circle" | "diamond" | "triangle" | "square" | "dash" | "dot" | "none"`
- `lineDataSymbolSize: number`
- `lineDataSymbolLineColor`, `lineDataSymbolLineSize` — marker stroke
- `lineDash: "solid" | "dash" | "dashDot" | "longDash" | ...`
- `trendlineType: "linear" | "log" | "exp" | "poly" | "movingAvg"` — add a trend line to a series

---

## Pie / Doughnut specific

- `holeSize: 10-90` — doughnut only; controls ring thickness
- `firstSliceAng: 0-359` — starting angle
- **No center text option.** Overlay an `addText` box at the chart's center manually.

---

## Radar specific

- `radarStyle: "standard" | "marker" | "filled"`
- 3-axis radars are always triangles — need 5+ axes for "radar" shape
- Filled mode respects `chartColorsOpacity`

**Limit:** all axes share one value scale. No per-axis min/max. For mixed-unit radars, normalize values to 0-100 per axis before passing.

---

## Combo chart specifics

Signature: `slide.addChart(dataArr, options)` where `dataArr = [{ type, data, options }, ...]`.

Each block's options can override series styling. Shared options (legend, plot area, axes arrays) go at the outer `options`.

Typical pattern: bar-type block with primary val axis, line-type block with `secondaryValAxis: true`.

---

## What this library does NOT support

**Chart types:** waterfall (native), Sankey, treemap, sunburst, heatmap, Gantt, funnel, box plot, candlestick, error bars (as first-class), violin, ridgeline.

**Styling:** per-point colors (bar / line / scatter), per-point data labels, gradient fills on series, pattern fills, conditional formatting, zone shading (e.g. "highlight values > 10"), per-axis scale on radar, center text on donut, chart title with per-run styling, per-series line dash, legend shape override.

**Layout:** small multiples (use multiple `addChart` calls manually), linked axes across charts, shared legends across charts.

**Data:** null handling in stacked bar is quirky; empty categories collapse rather than render as gaps.

**Interaction (in the rendered PPT):** none. Chart is static data embedded as `<c:chart>` XML. PowerPoint can edit it via the chart data pane; that's the only interactivity.

---

## When to escape to `claude-pptx-plot`

Escape to pure shapes when native fails structurally:

| Situation | Why native fails |
|---|---|
| Per-bar color needed | `chartColors` is series-level only |
| Waterfall with direction-colored deltas | No native waterfall + no per-point color |
| Bubble chart with correct absolute sizing | `bubbleScale` override isn't exposed |
| Complex quadrant / matrix charts | No native support; scatter can't tint background regions |
| Data viz on top of background shapes | Overlay alignment fragile; shape-based is cleaner |
| Story-driven annotation on specific points | Native has no per-point label control |

For everything else — start native.

---

## See also

- [pptxgenjs official docs](https://gitbrent.github.io/PptxGenJS/)
- [pptxgenjs demo gallery](https://gitbrent.github.io/PptxGenJS/demo/)
- `audit_deck.pptx` (in `output/`) — 14-slide walkthrough of the above at simple + stress-test complexity
