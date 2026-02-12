import path from "node:path";
import { createRequire } from "node:module";
import { THEME } from "./theme.js";
import {
  drawPageBg,
  drawTopHeader,
  drawLeftSidebar,
  drawCard,
  drawKpiCard,
  drawIconTextCard,
  drawTableCard,
  drawTimeline,
  txt,
  rect
} from "./components.js";

const require = createRequire(import.meta.url);
const PptxGenJS = require("pptxgenjs");

function blocksToMap(blocks = []) {
  const map = new Map();
  for (const b of blocks) map.set(b.heading, b.bullets);
  return map;
}

function resolveTemplate(s) {
  if (s.template) return s.template;

  const t = s.type;
  if (t === "title") return "title";
  if (t === "bullets") return "bulletsCard";
  if (t === "twoColumn") return "twoColumnCards";
  if (t === "comparisonTable") return "tableCard";
  if (t === "timeline") return "timelineSteps";
  if (t === "kpiCards") return "kpiAndPrinciples";

  return "bulletsCard";
}

/** ---------- Templates ---------- */

function renderTitle(slide, s) {
  drawPageBg(slide);
  rect(slide, { x: 0, y: 0, w: THEME.slideW, h: 2.3, fill: { color: THEME.primary }, line: { color: THEME.primary } });

  txt(slide, s.title || "Title", {
    x: THEME.margin, y: 0.85, w: THEME.slideW - THEME.margin * 2, h: 0.9,
    fontFace: THEME.font, fontSize: 40, bold: true, color: THEME.white
  });

  if (s.subtitle) {
    txt(slide, s.subtitle, {
      x: THEME.margin, y: 1.75, w: THEME.slideW - THEME.margin * 2, h: 0.4,
      fontFace: THEME.font, fontSize: 18, color: "DBEAFE"
    });
  }
}

/**
 * Minimum quality fallback: bullets always become a nice card.
 */
function renderBulletsCard(slide, s) {
  drawPageBg(slide);
  drawTopHeader(slide, { title: s.title || "Slide", iconChar: "•" });

  const items =
    Array.isArray(s.items) ? s.items :
    Array.isArray(s.blocks) ? s.blocks.flatMap((b) => b.bullets) :
    [];

  const body = items.map((t) => `• ${t}`).join("\n");

  slide.addShape("roundRect", {
    x: THEME.margin,
    y: THEME.topBarH + 0.6,
    w: THEME.slideW - THEME.margin * 2,
    h: THEME.slideH - (THEME.topBarH + 0.6) - THEME.margin,
    fill: { color: THEME.cardBg },
    line: { color: THEME.border, width: 1 },
    radius: THEME.radius
  });

  txt(slide, body || "• （内容なし）", {
    x: THEME.margin + 0.45,
    y: THEME.topBarH + 0.95,
    w: THEME.slideW - THEME.margin * 2 - 0.9,
    h: THEME.slideH - (THEME.topBarH + 0.95) - THEME.margin,
    fontFace: THEME.font,
    fontSize: 16,
    color: THEME.text,
    valign: "top",
    lineSpacingMultiple: 1.18
  });
}

/**
 * Left sidebar + vertical cards (your screenshot #2 style)
 */
function renderSidebarCards(slide, s) {
  drawPageBg(slide);
  const sidebarTitle = s.params?.sidebarTitle || s.title;
  const icon = s.params?.iconChar || "👥";
  const { sidebarW } = drawLeftSidebar(slide, { sidebarTitle, iconChar: icon, width: 4.2 });

  const x0 = sidebarW + THEME.gap;
  const w = THEME.slideW - x0 - THEME.margin;
  const cardH = 1.65;
  const yStart = 0.95;
  const yGap = 0.38;

  const bmap = blocksToMap(s.blocks);
  const order = s.params?.order
    ? String(s.params.order).split(",").map((t) => t.trim()).filter(Boolean)
    : Array.from(bmap.keys());

  const icons = { "定義": "i", "主な活動": "≡", "必要性": "!" };

  order.slice(0, 3).forEach((heading, i) => {
    const y = yStart + i * (cardH + yGap);
    drawCard(slide, {
      x: x0, y, w, h: cardH,
      heading,
      iconChar: icons[heading] || "●",
      bodyLines: (bmap.get(heading) || []).slice(0, 6)
    });
  });
}

/**
 * Header + grid cards (3 + 2)
 */
function renderCaseCardsGrid(slide, s) {
  drawPageBg(slide);
  drawTopHeader(slide, { title: s.title, iconChar: "▦" });

  const cards = (s.blocks || []).map((b) => ({ heading: b.heading, lines: b.bullets }));
  const x0 = THEME.margin;
  const y0 = THEME.topBarH + 0.55;
  const wArea = THEME.slideW - THEME.margin * 2;
  const gap = THEME.gap;

  const topCols = 3;
  const cardWTop = (wArea - gap * (topCols - 1)) / topCols;
  const cardHTop = 2.15;

  for (let i = 0; i < Math.min(3, cards.length); i++) {
    const c = cards[i];
    const x = x0 + i * (cardWTop + gap);
    drawCard(slide, {
      x, y: y0, w: cardWTop, h: cardHTop,
      heading: c.heading,
      iconChar: "■",
      bodyLines: (c.lines || []).slice(0, 5)
    });
  }

  const bottom = cards.slice(3, 5);
  if (bottom.length) {
    const bottomCols = 2;
    const cardWBot = (wArea - gap * (bottomCols - 1)) / bottomCols;
    const cardHBot = 2.15;
    const y = y0 + cardHTop + 0.55;

    for (let i = 0; i < bottom.length; i++) {
      const c = bottom[i];
      const x = x0 + i * (cardWBot + gap);
      drawCard(slide, {
        x, y, w: cardWBot, h: cardHBot,
        heading: c.heading,
        iconChar: "■",
        bodyLines: (c.lines || []).slice(0, 5)
      });
    }
  }
}

/**
 * KPI row + principles row
 */
function renderKpiAndPrinciples(slide, s) {
  drawPageBg(slide);
  drawTopHeader(slide, { title: s.title, iconChar: "≣" });

  const bmap = blocksToMap(s.blocks);
  const kpi = bmap.get("KPI") || [];
  const points = bmap.get("成功のポイント") || bmap.get("ポイント") || [];

  const y0 = THEME.topBarH + 0.7;

  const kpiCards = kpi.slice(0, 3).map((line) => {
    const m = String(line).match(/^([^:]+)\s*:\s*(.+)$/);
    return m ? { value: m[1].trim(), label: m[2].trim() } : { value: line, label: "" };
  });

  const gap = THEME.gap;
  const wArea = THEME.slideW - THEME.margin * 2;
  const kW = (wArea - gap * 2) / 3;

  for (let i = 0; i < kpiCards.length; i++) {
    drawKpiCard(slide, {
      x: THEME.margin + i * (kW + gap),
      y: y0,
      w: kW,
      h: 2.0,
      value: kpiCards[i].value,
      label: kpiCards[i].label
    });
  }

  txt(slide, "成功のポイント", {
    x: THEME.margin, y: y0 + 2.35, w: 5, h: 0.3,
    fontFace: THEME.font, fontSize: 12, bold: true, color: THEME.muted
  });

  const bottomY = y0 + 2.7;
  const cols = 5;
  const cW = (wArea - gap * (cols - 1)) / cols;

  const iconCycle = ["★", "⚙", "🎓", "👥", "🛡"];
  for (let i = 0; i < Math.min(cols, points.length); i++) {
    const title = points[i];
    drawIconTextCard(slide, {
      x: THEME.margin + i * (cW + gap),
      y: bottomY,
      w: cW,
      h: 2.2,
      title,
      text: "",
      iconChar: iconCycle[i % iconCycle.length]
    });
  }
}

/**
 * NEW: Two-column cards template
 * - Left card: heading + bullets
 * - Right card: heading + bullets
 */
function renderTwoColumnCards(slide, s) {
  drawPageBg(slide);
  drawTopHeader(slide, { title: s.title || "Two Column", iconChar: "▤" });

  const y = THEME.topBarH + 0.65;
  const h = THEME.slideH - y - THEME.margin;
  const gap = THEME.gap;
  const wArea = THEME.slideW - THEME.margin * 2;
  const w = (wArea - gap) / 2;

  // spec may provide left/right already (md_to_spec for type=twoColumn)
  const left = s.left || (s.blocks?.[0] ? { heading: s.blocks[0].heading, bullets: s.blocks[0].bullets } : { heading: "左", bullets: [] });
  const right = s.right || (s.blocks?.[1] ? { heading: s.blocks[1].heading, bullets: s.blocks[1].bullets } : { heading: "右", bullets: [] });

  drawCard(slide, {
    x: THEME.margin,
    y,
    w,
    h,
    heading: left.heading || "左",
    iconChar: "◧",
    bodyLines: (left.bullets || []).slice(0, 10)
  });

  drawCard(slide, {
    x: THEME.margin + w + gap,
    y,
    w,
    h,
    heading: right.heading || "右",
    iconChar: "◨",
    bodyLines: (right.bullets || []).slice(0, 10)
  });
}

/**
 * NEW: Table card template (comparison)
 */
function renderTableCard(slide, s) {
  drawPageBg(slide);
  drawTopHeader(slide, { title: s.title || "Table", iconChar: "▦" });

  const x = THEME.margin;
  const y = THEME.topBarH + 0.65;
  const w = THEME.slideW - THEME.margin * 2;
  const h = THEME.slideH - y - THEME.margin;

  // spec may provide headers/rows (md_to_spec for type=comparisonTable)
  const headers = s.headers || [];
  const rows = s.rows || [];

  drawTableCard(slide, {
    x, y, w, h,
    title: s.params?.tableTitle || "比較表",
    iconChar: "▦",
    headers,
    rows
  });
}

/**
 * NEW: Timeline steps template
 */
function renderTimelineSteps(slide, s) {
  drawPageBg(slide);
  drawTopHeader(slide, { title: s.title || "Timeline", iconChar: "⟷" });

  const x = THEME.margin;
  const y = THEME.topBarH + 0.65;
  const w = THEME.slideW - THEME.margin * 2;
  const h = THEME.slideH - y - THEME.margin;

  // spec may provide steps (md_to_spec for type=timeline)
  const steps =
    Array.isArray(s.steps) ? s.steps :
    Array.isArray(s.blocks) ? s.blocks.map((b) => ({ title: b.heading, body: b.bullets })) :
    [];

  drawTimeline(slide, {
    x, y, w, h,
    title: s.params?.timelineTitle || "タイムライン",
    steps: steps.slice(0, 5) // 5ステップ程度が見栄え安定
  });
}

export async function generatePptx(spec, outPath) {
  const pptx = new PptxGenJS();
  pptx.layout = spec.layout || "LAYOUT_WIDE";
  if (spec.theme?.fontFace) pptx.theme = { headFontFace: spec.theme.fontFace, bodyFontFace: spec.theme.fontFace };

  for (const s of spec.slides || []) {
    const slide = pptx.addSlide();
    const tpl = resolveTemplate(s);

    if (tpl === "title") renderTitle(slide, s);
    else if (tpl === "sidebarCards") renderSidebarCards(slide, s);
    else if (tpl === "caseCardsGrid") renderCaseCardsGrid(slide, s);
    else if (tpl === "kpiAndPrinciples") renderKpiAndPrinciples(slide, s);
    else if (tpl === "bulletsCard") renderBulletsCard(slide, s);

    // NEW templates
    else if (tpl === "twoColumnCards") renderTwoColumnCards(slide, s);
    else if (tpl === "tableCard") renderTableCard(slide, s);
    else if (tpl === "timelineSteps") renderTimelineSteps(slide, s);

    else renderBulletsCard(slide, s);
  }

  const outAbs = path.resolve(process.cwd(), outPath);
  await pptx.writeFile({ fileName: outAbs });
  return outAbs;
}

// CLI debug
if (process.argv[1]?.endsWith("generate.mjs")) {
  const fs = await import("node:fs");
  const [, , specPath, outPptx] = process.argv;
  if (!specPath) {
    console.error("Usage: node generate.mjs <spec.json> <out.pptx>");
    process.exit(1);
  }
  const spec = JSON.parse(fs.readFileSync(specPath, "utf-8"));
  await generatePptx(spec, outPptx || "output.pptx");
  console.log("✅ done");
}
