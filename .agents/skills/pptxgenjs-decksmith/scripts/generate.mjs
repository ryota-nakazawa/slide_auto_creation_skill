import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const PptxGenJS = require("pptxgenjs");

function mustString(v, name) {
  if (typeof v !== "string" || !v.trim()) throw new Error(`Missing string: ${name}`);
  return v;
}
function asString(v, fallback = "") {
  if (v === null || v === undefined) return fallback;
  return String(v);
}

// Layout constants (16:9)
const M = 0.6;
const W = 13.33;
const H = 7.5;
const TITLE_Y = 0.35;
const TITLE_H = 0.6;
const BODY_Y = 1.25;
const BODY_H = 5.9;

function addTopBar(slide, title) {
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.25, fill: { color: "F3F4F6" }, line: { color: "F3F4F6" } });
  slide.addText(title, { x: M, y: TITLE_Y, w: W - 2 * M, h: TITLE_H, fontSize: 28, bold: true, color: "111827" });
}

function addTitle(slide, spec) {
  const title = mustString(spec.title, "slides[].title");
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 2.2, fill: { color: "111827" }, line: { color: "111827" } });
  slide.addText(title, { x: M, y: 0.75, w: W - 2 * M, h: 0.9, fontSize: 42, bold: true, color: "FFFFFF" });
  if (spec.subtitle) slide.addText(asString(spec.subtitle), { x: M, y: 1.65, w: W - 2 * M, h: 0.5, fontSize: 18, color: "D1D5DB" });
  slide.addText("Generated with PptxGenJS", { x: M, y: H - 0.55, w: W - 2 * M, h: 0.3, fontSize: 10, color: "6B7280" });
}

function addBullets(slide, spec) {
  const title = mustString(spec.title, "slides[].title");
  addTopBar(slide, title);

  const items = Array.isArray(spec.items) ? spec.items.map(String) : [];
  const bulletText = items.map((t) => `• ${t}`).join("\n");

  slide.addShape("roundRect", { x: M, y: BODY_Y, w: W - 2 * M, h: BODY_H, fill: { color: "FFFFFF" }, line: { color: "E5E7EB" }, radius: 0.15 });
  slide.addText(bulletText, { x: M + 0.35, y: BODY_Y + 0.35, w: W - 2 * M - 0.7, h: BODY_H - 0.7, fontSize: 18, valign: "top", color: "111827" });
}

function addTwoColumn(slide, spec) {
  const title = mustString(spec.title, "slides[].title");
  addTopBar(slide, title);

  const left = spec.left || {};
  const right = spec.right || {};
  const colGap = 0.35;
  const colW = (W - 2 * M - colGap) / 2;
  const cardY = BODY_Y;
  const cardH = BODY_H;

  slide.addShape("roundRect", { x: M, y: cardY, w: colW, h: cardH, fill: { color: "FFFFFF" }, line: { color: "E5E7EB" }, radius: 0.15 });
  slide.addText(asString(left.heading, "左カラム"), { x: M + 0.3, y: cardY + 0.25, w: colW - 0.6, h: 0.35, fontSize: 16, bold: true, color: "111827" });

  const leftItems = Array.isArray(left.items) ? left.items.map(String) : [];
  const leftText = leftItems.length ? leftItems.map((t) => `• ${t}`).join("\n") : asString(left.text, "");
  slide.addText(leftText, { x: M + 0.3, y: cardY + 0.75, w: colW - 0.6, h: cardH - 1.05, fontSize: 16, color: "111827", valign: "top" });

  const rx = M + colW + colGap;
  slide.addShape("roundRect", { x: rx, y: cardY, w: colW, h: cardH, fill: { color: "FFFFFF" }, line: { color: "E5E7EB" }, radius: 0.15 });
  slide.addText(asString(right.heading, "右カラム"), { x: rx + 0.3, y: cardY + 0.25, w: colW - 0.6, h: 0.35, fontSize: 16, bold: true, color: "111827" });

  // allow either items or text or image
  if (right.imagePath) {
    slide.addImage({ path: String(right.imagePath), x: rx + 0.3, y: cardY + 0.75, w: colW - 0.6, h: cardH - 1.05 });
  } else {
    const rightItems = Array.isArray(right.items) ? right.items.map(String) : [];
    const rightText = rightItems.length ? rightItems.map((t) => `• ${t}`).join("\n") : asString(right.text, "");
    slide.addText(rightText, { x: rx + 0.3, y: cardY + 0.75, w: colW - 0.6, h: cardH - 1.05, fontSize: 16, color: "111827", valign: "top" });
  }
}

function addKpiCards(slide, spec) {
  const title = mustString(spec.title, "slides[].title");
  addTopBar(slide, title);

  const cards = Array.isArray(spec.cards) ? spec.cards : [];
  if (cards.length === 0) {
    slide.addText("cards が空です", { x: M, y: BODY_Y, w: W - 2 * M, h: 1, fontSize: 16, color: "B91C1C" });
    return;
  }

  const gridCols = cards.length <= 2 ? 2 : 2;
  const gridRows = Math.ceil(cards.length / gridCols);
  const gap = 0.35;
  const cardW = (W - 2 * M - gap * (gridCols - 1)) / gridCols;
  const cardH = Math.min(2.3, (BODY_H - gap * (gridRows - 1)) / gridRows);

  let idx = 0;
  for (let r = 0; r < gridRows; r++) {
    for (let c = 0; c < gridCols; c++) {
      if (idx >= cards.length) break;
      const item = cards[idx++];
      const x = M + c * (cardW + gap);
      const y = BODY_Y + r * (cardH + gap);

      slide.addShape("roundRect", { x, y, w: cardW, h: cardH, fill: { color: "FFFFFF" }, line: { color: "E5E7EB" }, radius: 0.2 });
      slide.addShape("rect", { x, y, w: cardW, h: 0.18, fill: { color: "111827" }, line: { color: "111827" } });

      slide.addText(asString(item.label, ""), { x: x + 0.35, y: y + 0.35, w: cardW - 0.7, h: 0.4, fontSize: 14, color: "6B7280" });
      slide.addText(asString(item.value, ""), { x: x + 0.35, y: y + 0.85, w: cardW - 0.7, h: 0.8, fontSize: 30, bold: true, color: "111827" });
      if (item.note) slide.addText(asString(item.note), { x: x + 0.35, y: y + 1.65, w: cardW - 0.7, h: 0.4, fontSize: 12, color: "6B7280" });
    }
  }
}

function addComparisonTable(slide, spec) {
  const title = mustString(spec.title, "slides[].title");
  addTopBar(slide, title);

  const columns = Array.isArray(spec.columns) ? spec.columns.map(String) : [];
  const rows = Array.isArray(spec.rows) ? spec.rows.map((r) => (Array.isArray(r) ? r.map((v) => asString(v)) : [])) : [];

  slide.addShape("roundRect", { x: M, y: BODY_Y, w: W - 2 * M, h: BODY_H, fill: { color: "FFFFFF" }, line: { color: "E5E7EB" }, radius: 0.15 });

  const data = [];
  if (columns.length) data.push(columns);
  data.push(...rows);

  slide.addTable(data, {
    x: M + 0.25,
    y: BODY_Y + 0.35,
    w: W - 2 * M - 0.5,
    h: BODY_H - 0.7,
    fontSize: 14,
    border: { pt: 1, color: "E5E7EB" },
    fill: "FFFFFF",
    color: "111827",
    valign: "middle",
    rowH: 0.35
  });
}

function addImageHero(slide, spec) {
  const title = mustString(spec.title, "slides[].title");
  const img = mustString(spec.imagePath, "slides[].imagePath");

  slide.addImage({ path: img, x: 0, y: 0, w: W, h: H });
  slide.addShape("rect", { x: 0, y: H - 1.35, w: W, h: 1.35, fill: { color: "111827" }, line: { color: "111827" } });

  slide.addText(title, { x: M, y: H - 1.15, w: W - 2 * M, h: 0.55, fontSize: 34, bold: true, color: "FFFFFF" });
  if (spec.subtitle) slide.addText(asString(spec.subtitle), { x: M, y: H - 0.60, w: W - 2 * M, h: 0.35, fontSize: 16, color: "D1D5DB" });
  if (spec.caption) slide.addText(asString(spec.caption), { x: M, y: H - 0.30, w: W - 2 * M, h: 0.25, fontSize: 10, color: "9CA3AF" });
}

function addTimeline(slide, spec) {
  const title = mustString(spec.title, "slides[].title");
  addTopBar(slide, title);

  const items = Array.isArray(spec.items) ? spec.items : [];
  if (items.length < 2) {
    slide.addText("timeline は items を2つ以上入れてください", { x: M, y: BODY_Y, w: W - 2 * M, h: 1, fontSize: 16, color: "B91C1C" });
    return;
  }

  const lineY = BODY_Y + 2.2;
  slide.addShape("line", { x: M + 0.4, y: lineY, w: W - 2 * M - 0.8, h: 0, line: { color: "9CA3AF", width: 2 } });

  const span = W - 2 * M - 0.8;
  const step = span / (items.length - 1);

  items.forEach((it, i) => {
    const x = M + 0.4 + step * i;
    slide.addShape("ellipse", { x: x - 0.14, y: lineY - 0.14, w: 0.28, h: 0.28, fill: { color: "111827" }, line: { color: "111827" } });
    slide.addText(asString(it.label, `Step ${i + 1}`), { x: x - 1.0, y: lineY - 0.85, w: 2.0, h: 0.3, fontSize: 12, bold: true, color: "111827", align: "center" });
    slide.addText(asString(it.text, ""), { x: x - 1.2, y: lineY + 0.25, w: 2.4, h: 1.4, fontSize: 12, color: "374151", align: "center", valign: "top" });
  });
}

function addChart(slide, spec, pptx) {
  const title = mustString(spec.title, "slides[].title");
  addTopBar(slide, title);

  const chartTypeStr = asString(spec.chartType, "bar").toLowerCase();
  const chartType = pptx.ChartType?.[chartTypeStr];

  if (!chartType) {
    slide.addText(`Unsupported chartType: ${chartTypeStr}`, { x: M, y: BODY_Y, w: W - 2 * M, h: 1, fontSize: 16, color: "B91C1C" });
    slide.addText("Allowed: area, bar, bar3d, bubble, bubble3d, doughnut, line, pie, radar, scatter", { x: M, y: BODY_Y + 0.6, w: W - 2 * M, h: 1, fontSize: 12, color: "6B7280" });
    return;
  }

  const series = Array.isArray(spec.series) ? spec.series : [];
  if (series.length === 0) {
    slide.addText("chart は series を1つ以上入れてください", { x: M, y: BODY_Y, w: W - 2 * M, h: 1, fontSize: 16, color: "B91C1C" });
    return;
  }

  slide.addShape("roundRect", { x: M, y: BODY_Y, w: W - 2 * M, h: BODY_H, fill: { color: "FFFFFF" }, line: { color: "E5E7EB" }, radius: 0.15 });
  slide.addChart(chartType, series, { x: M + 0.35, y: BODY_Y + 0.35, w: W - 2 * M - 0.7, h: BODY_H - 0.7, ...(spec.options || {}) });
}

export async function generatePptx(spec, outPath) {
  const pptx = new PptxGenJS();
  pptx.layout = spec.layout || "LAYOUT_WIDE";
  if (spec.theme?.fontFace) pptx.theme = { headFontFace: spec.theme.fontFace, bodyFontFace: spec.theme.fontFace };

  for (const s of spec.slides || []) {
    const slide = pptx.addSlide();
    const type = s.type;

    if (type === "title") addTitle(slide, s);
    else if (type === "imageHero") addImageHero(slide, s);
    else if (type === "bullets") addBullets(slide, s);
    else if (type === "twoColumn") addTwoColumn(slide, s);
    else if (type === "kpiCards") addKpiCards(slide, s);
    else if (type === "comparisonTable") addComparisonTable(slide, s);
    else if (type === "timeline") addTimeline(slide, s);
    else if (type === "chart") addChart(slide, s, pptx);
    else {
      slide.addText(`Unsupported slide type: ${String(type)}`, { x: 0.7, y: 0.7, fontSize: 20 });
      slide.addText(JSON.stringify(s, null, 2), { x: 0.7, y: 1.5, w: 12, h: 5, fontSize: 10 });
    }
  }

  const outAbs = path.resolve(process.cwd(), outPath);
  await pptx.writeFile({ fileName: outAbs });
  return outAbs;
}

// CLI (debug)
if (process.argv[1]?.endsWith("generate.mjs")) {
  const [, , specPath, outPptx] = process.argv;
  if (!specPath) {
    console.error("Usage: node generate.mjs <spec.json> <out.pptx>");
    process.exit(1);
  }
  const fs = await import("node:fs");
  const spec = JSON.parse(fs.readFileSync(specPath, "utf-8"));
  await generatePptx(spec, outPptx || "output.pptx");
  console.log("✅ done");
}
