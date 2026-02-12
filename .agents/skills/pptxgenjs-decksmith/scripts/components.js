import { THEME } from "./theme.js";

// Utility
export function txt(slide, text, opt) {
  slide.addText(String(text ?? ""), opt);
}
export function rect(slide, opt) {
  slide.addShape("rect", opt);
}
export function roundRect(slide, opt) {
  slide.addShape("roundRect", opt);
}

function shapeWithOptionalShadow(base) {
  return {
    ...base,
    ...(THEME.shadow ? { shadow: THEME.shadow } : {})
  };
}

export function iconBadge(slide, { x, y, textChar }) {
  // Simple “icon” as a rounded badge (emoji/1char). Replace with SVG/PNG later if desired.
  slide.addShape("roundRect", shapeWithOptionalShadow({
    x, y, w: 0.42, h: 0.42,
    fill: { color: THEME.primary2 },
    line: { color: THEME.primary2 },
    radius: 0.12
  }));

  txt(slide, textChar || "●", {
    x, y: y + 0.02, w: 0.42, h: 0.42,
    fontFace: THEME.font, fontSize: 14, color: THEME.white,
    align: "center", valign: "mid"
  });
}

// Background
export function drawPageBg(slide) {
  rect(slide, { x: 0, y: 0, w: THEME.slideW, h: THEME.slideH, fill: { color: THEME.bg }, line: { color: THEME.bg } });
}

// Top header band
export function drawTopHeader(slide, { title, iconChar = "■" }) {
  rect(slide, { x: 0, y: 0, w: THEME.slideW, h: THEME.topBarH, fill: { color: THEME.primary }, line: { color: THEME.primary } });

  // icon
  iconBadge(slide, { x: THEME.margin, y: 0.17, textChar: iconChar });

  // title
  txt(slide, title, {
    x: THEME.margin + 0.55,
    y: 0.14,
    w: THEME.slideW - THEME.margin * 2 - 0.55,
    h: 0.55,
    fontFace: THEME.font,
    fontSize: THEME.titleSize,
    bold: true,
    color: THEME.white,
    valign: "mid"
  });
}

// Left sidebar (slide style)
export function drawLeftSidebar(slide, { sidebarTitle, iconChar = "●", width = 4.2 }) {
  rect(slide, { x: 0, y: 0, w: width, h: THEME.slideH, fill: { color: THEME.primary }, line: { color: THEME.primary } });

  // Large icon cluster (placeholder)
  iconBadge(slide, { x: 1.3, y: 2.55, textChar: iconChar });
  iconBadge(slide, { x: 1.85, y: 2.55, textChar: iconChar });
  iconBadge(slide, { x: 1.57, y: 3.05, textChar: iconChar });

  txt(slide, sidebarTitle, {
    x: 0.6,
    y: 3.6,
    w: width - 1.2,
    h: 1.6,
    fontFace: THEME.font,
    fontSize: 26,
    bold: true,
    color: THEME.white,
    align: "center",
    valign: "mid"
  });

  return { sidebarW: width };
}

// Generic card
export function drawCard(slide, { x, y, w, h, heading, iconChar, bodyLines }) {
  slide.addShape("roundRect", shapeWithOptionalShadow({
    x, y, w, h,
    fill: { color: THEME.cardBg },
    line: { color: THEME.border, width: 1 },
    radius: THEME.radius
  }));

  const pad = THEME.cardPad;
  const headY = y + pad;

  if (iconChar) iconBadge(slide, { x: x + pad, y: headY + 0.02, textChar: iconChar });

  txt(slide, heading || "", {
    x: x + pad + (iconChar ? 0.55 : 0),
    y: headY,
    w: w - pad * 2 - (iconChar ? 0.55 : 0),
    h: 0.35,
    fontFace: THEME.font,
    fontSize: THEME.h2Size,
    bold: true,
    color: THEME.text,
    valign: "mid"
  });

  const body = (bodyLines || []).map((t) => `• ${t}`).join("\n");
  txt(slide, body, {
    x: x + pad,
    y: headY + 0.45,
    w: w - pad * 2,
    h: h - pad * 2 - 0.45,
    fontFace: THEME.font,
    fontSize: THEME.bodySize,
    color: THEME.text,
    valign: "top",
    lineSpacingMultiple: 1.15
  });
}

// KPI card
export function drawKpiCard(slide, { x, y, w, h, value, label }) {
  slide.addShape("roundRect", shapeWithOptionalShadow({
    x, y, w, h,
    fill: { color: THEME.cardBg },
    line: { color: THEME.border, width: 1 },
    radius: THEME.radius
  }));

  // top accent
  rect(slide, { x, y, w, h: 0.10, fill: { color: THEME.primary2 }, line: { color: THEME.primary2 } });

  txt(slide, value, {
    x, y: y + 0.45,
    w, h: 0.7,
    fontFace: THEME.font,
    fontSize: THEME.kpiValueSize,
    bold: true,
    color: THEME.primary,
    align: "center",
    valign: "mid"
  });

  txt(slide, label, {
    x, y: y + 1.25,
    w, h: 0.35,
    fontFace: THEME.font,
    fontSize: THEME.kpiLabelSize,
    color: THEME.muted,
    align: "center",
    valign: "mid"
  });
}

// Small icon card
export function drawIconTextCard(slide, { x, y, w, h, title, text, iconChar = "●" }) {
  slide.addShape("roundRect", shapeWithOptionalShadow({
    x, y, w, h,
    fill: { color: THEME.cardBg },
    line: { color: THEME.border, width: 1 },
    radius: THEME.radius
  }));

  iconBadge(slide, { x: x + THEME.cardPad, y: y + THEME.cardPad, textChar: iconChar });

  txt(slide, title, {
    x: x + THEME.cardPad + 0.55,
    y: y + THEME.cardPad - 0.02,
    w: w - THEME.cardPad * 2 - 0.55,
    h: 0.35,
    fontFace: THEME.font,
    fontSize: THEME.h2Size,
    bold: true,
    color: THEME.text
  });

  txt(slide, text, {
    x: x + THEME.cardPad,
    y: y + THEME.cardPad + 0.55,
    w: w - THEME.cardPad * 2,
    h: h - THEME.cardPad * 2 - 0.55,
    fontFace: THEME.font,
    fontSize: 11,
    color: THEME.muted,
    valign: "top",
    lineSpacingMultiple: 1.15
  });
}

/**
 * Table (simple grid) inside a card.
 * headers: string[]
 * rows: string[][]
 */
export function drawTableCard(slide, { x, y, w, h, title, iconChar = "▦", headers = [], rows = [] }) {
  slide.addShape("roundRect", shapeWithOptionalShadow({
    x, y, w, h,
    fill: { color: THEME.cardBg },
    line: { color: THEME.border, width: 1 },
    radius: THEME.radius
  }));

  const pad = THEME.cardPad;
  const headY = y + pad;

  iconBadge(slide, { x: x + pad, y: headY + 0.02, textChar: iconChar });
  txt(slide, title || "", {
    x: x + pad + 0.55,
    y: headY,
    w: w - pad * 2 - 0.55,
    h: 0.35,
    fontFace: THEME.font,
    fontSize: THEME.h2Size,
    bold: true,
    color: THEME.text,
    valign: "mid"
  });

  const gridX = x + pad;
  const gridY = headY + 0.55;
  const gridW = w - pad * 2;
  const gridH = h - pad * 2 - 0.55;

  // Grid sizing
  const cols = Math.max(1, headers.length || (rows[0]?.length ?? 1));
  const colW = gridW / cols;

  const headerH = 0.38;
  const rowH = Math.min(0.42, (gridH - headerH) / Math.max(1, rows.length));

  // Header background
  rect(slide, { x: gridX, y: gridY, w: gridW, h: headerH, fill: { color: "EFF6FF" }, line: { color: THEME.border, width: 1 } });

  // Header text + vertical lines
  for (let c = 0; c < cols; c++) {
    const hx = gridX + c * colW;
    if (c > 0) rect(slide, { x: hx, y: gridY, w: 0.001, h: headerH + rowH * rows.length, fill: { color: THEME.border }, line: { color: THEME.border } });
    txt(slide, headers[c] ?? "", {
      x: hx + 0.08,
      y: gridY + 0.05,
      w: colW - 0.16,
      h: headerH - 0.1,
      fontFace: THEME.font,
      fontSize: 11,
      bold: true,
      color: THEME.primary,
      valign: "mid"
    });
  }

  // Horizontal line under header
  rect(slide, { x: gridX, y: gridY + headerH, w: gridW, h: 0.001, fill: { color: THEME.border }, line: { color: THEME.border } });

  // Rows
  for (let r = 0; r < rows.length; r++) {
    const ry = gridY + headerH + r * rowH;

    // zebra background (very subtle)
    if (r % 2 === 1) rect(slide, { x: gridX, y: ry, w: gridW, h: rowH, fill: { color: "FBFDFF" }, line: { color: "FBFDFF" } });

    // row bottom line
    rect(slide, { x: gridX, y: ry + rowH, w: gridW, h: 0.001, fill: { color: THEME.border }, line: { color: THEME.border } });

    for (let c = 0; c < cols; c++) {
      const cx = gridX + c * colW;
      const cell = (rows[r]?.[c] ?? "").toString();
      txt(slide, cell, {
        x: cx + 0.08,
        y: ry + 0.06,
        w: colW - 0.16,
        h: rowH - 0.12,
        fontFace: THEME.font,
        fontSize: 11,
        color: THEME.text,
        valign: "mid"
      });
    }
  }
}

/**
 * Timeline steps (horizontal)
 * steps: { title: string, body: string[] }[]
 */
export function drawTimeline(slide, { x, y, w, h, title, steps = [] }) {
  slide.addShape("roundRect", shapeWithOptionalShadow({
    x, y, w, h,
    fill: { color: THEME.cardBg },
    line: { color: THEME.border, width: 1 },
    radius: THEME.radius
  }));

  const pad = THEME.cardPad;
  iconBadge(slide, { x: x + pad, y: y + pad + 0.02, textChar: "⟷" });
  txt(slide, title || "", {
    x: x + pad + 0.55,
    y: y + pad,
    w: w - pad * 2 - 0.55,
    h: 0.35,
    fontFace: THEME.font,
    fontSize: THEME.h2Size,
    bold: true,
    color: THEME.text,
    valign: "mid"
  });

  const areaX = x + pad;
  const areaY = y + pad + 0.6;
  const areaW = w - pad * 2;
  const areaH = h - pad * 2 - 0.6;

  const n = Math.max(1, steps.length);
  const stepW = areaW / n;
  const lineY = areaY + 0.55;

  // baseline
  rect(slide, { x: areaX + 0.25, y: lineY, w: areaW - 0.5, h: 0.04, fill: { color: "DBEAFE" }, line: { color: "DBEAFE" } });

  for (let i = 0; i < steps.length; i++) {
    const sx = areaX + i * stepW;

    // node
    slide.addShape("roundRect", {
      x: sx + stepW / 2 - 0.16,
      y: lineY - 0.16,
      w: 0.32,
      h: 0.32,
      fill: { color: THEME.primary2 },
      line: { color: THEME.primary2 },
      radius: 0.16
    });

    // step title
    txt(slide, steps[i].title || `Step ${i + 1}`, {
      x: sx + 0.1,
      y: areaY,
      w: stepW - 0.2,
      h: 0.35,
      fontFace: THEME.font,
      fontSize: 12,
      bold: true,
      color: THEME.primary,
      align: "center"
    });

    // body
    const body = (steps[i].body || []).map((t) => `• ${t}`).join("\n");
    txt(slide, body, {
      x: sx + 0.1,
      y: lineY + 0.25,
      w: stepW - 0.2,
      h: areaH - 0.9,
      fontFace: THEME.font,
      fontSize: 11,
      color: THEME.text,
      valign: "top",
      lineSpacingMultiple: 1.15
    });
  }
}
