// Slide chrome — eyebrows, titles, hairlines, logo, footers.
// These are the visual vocabulary of a Saperia slide, independent of charts.
// Lifted from handoff/build_deck.js and generalized to take (slide, theme).

function bgFill(slide, theme) {
  slide.background = { color: theme.colors.BG };
}

function addEyebrow(slide, theme, text, x, y, w) {
  slide.addText(text, {
    x, y, w: w || 8, h: 0.28,
    fontFace: theme.fonts.SANS,
    fontSize: 10,
    bold: true,
    color: theme.colors.MUTED,
    charSpacing: 2,
    valign: "top",
    margin: 0,
  });
}

function addTitle(slide, theme, text, opts = {}) {
  const { x = theme.layout.MARGIN, y = 0.82, w, fontSize = 30 } = opts;
  const width = w || (theme.layout.SLIDE_W - 2 * theme.layout.MARGIN);
  slide.addText(text, {
    x, y, w: width, h: 0.7,
    fontFace: theme.fonts.DISPLAY,
    fontSize,
    color: theme.colors.INK,
    margin: 0,
    valign: "top",
  });
}

function addSubtitle(slide, theme, text, opts = {}) {
  const { x = theme.layout.MARGIN, y = 1.45, w, fontSize = 13 } = opts;
  const width = w || (theme.layout.SLIDE_W - 2 * theme.layout.MARGIN);
  slide.addText(text, {
    x, y, w: width, h: 0.32,
    fontFace: theme.fonts.DISPLAY,
    fontSize,
    italic: true,
    color: theme.colors.INK,
    margin: 0,
    valign: "top",
  });
}

function addHairline(slide, theme, y, opts = {}) {
  const { x, w, color } = opts;
  const contentX = x ?? theme.layout.MARGIN;
  const contentW = w ?? (theme.layout.SLIDE_W - 2 * theme.layout.MARGIN);
  slide.addShape("line", {
    x: contentX,
    y,
    w: contentW,
    h: 0,
    line: { color: color || theme.colors.RULE, width: 0.5 },
  });
}

function addInkUnderscore(slide, theme, x, y, w) {
  slide.addShape("line", {
    x, y, w, h: 0,
    line: { color: theme.colors.INK, width: 1.5 },
  });
}

function addVerticalHairline(slide, theme, x, y, h) {
  slide.addShape("line", {
    x, y, w: 0, h,
    line: { color: theme.colors.RULE, width: 0.5 },
  });
}

function addLogo(slide, theme, logoPath) {
  if (!logoPath) return;
  const { SLIDE_W, SLIDE_H, MARGIN, LOGO_W, LOGO_H } = theme.layout;
  slide.addImage({
    path: logoPath,
    x: SLIDE_W - MARGIN - LOGO_W,
    y: SLIDE_H - MARGIN - LOGO_H + 0.1,
    w: LOGO_W,
    h: LOGO_H,
  });
}

function addSource(slide, theme, text) {
  const { SLIDE_W, SLIDE_H, MARGIN, LOGO_W } = theme.layout;
  slide.addText(text, {
    x: MARGIN,
    y: SLIDE_H - MARGIN - 0.2,
    w: SLIDE_W - 2 * MARGIN - LOGO_W - 0.3,
    h: 0.22,
    fontFace: theme.fonts.SANS,
    fontSize: 8,
    color: theme.colors.MUTED,
    italic: true,
    valign: "bottom",
    margin: 0,
  });
}

module.exports = {
  bgFill,
  addEyebrow,
  addTitle,
  addSubtitle,
  addHairline,
  addInkUnderscore,
  addVerticalHairline,
  addLogo,
  addSource,
};
