/**
 * slides.md format (recommended):
 * - H2 ("##") begins a new slide section
 * - Optional HTML comment hint: <!-- slide:type=twoColumn ... -->
 * - Content:
 *   - bullets: "- item"
 *   - title slide: "- title: X" "- subtitle: Y"
 *   - kpiCards: "- 売上: 100" style lines
 *   - timeline: "- Label: Text" style lines
 *   - comparisonTable: markdown table or "columns=" hint
 *
 * This parser is intentionally simple and robust.
 */

function parseFrontmatter(md) {
  const m = md.match(/^---\s*\n([\s\S]*?)\n---\s*\n/);
  if (!m) return { deck: {}, body: md };
  const fm = m[1];
  const body = md.slice(m[0].length);
  const deck = {};
  // super simple YAML-ish parse for deck keys
  // supports:
  // deck:
  //   layout: LAYOUT_WIDE
  //   themeFont: Aptos
  const lines = fm.split(/\r?\n/);
  let inDeck = false;
  for (const line of lines) {
    if (line.trim() === "deck:") {
      inDeck = true;
      continue;
    }
    if (inDeck) {
      const mm = line.match(/^\s+([A-Za-z0-9_]+)\s*:\s*(.+)\s*$/);
      if (mm) deck[mm[1]] = mm[2].trim();
    }
  }
  return { deck, body };
}

function parseHint(line) {
  // <!-- slide:type=timeline foo=bar -->
  const m = line.match(/<!--\s*slide:type=([a-zA-Z0-9_]+)([\s\S]*?)-->/);
  if (!m) return null;
  const type = m[1].trim();
  const rest = (m[2] || "").trim();

  const params = {};
  // parse key=value tokens
  // supports key=value and key="value with spaces"
  const re = /([a-zA-Z0-9_]+)\s*=\s*("([^"]+)"|[^\s]+)/g;
  let mm;
  while ((mm = re.exec(rest))) {
    const key = mm[1];
    const raw = mm[3] ?? mm[2];
    params[key] = String(raw).replace(/^"|"$/g, "");
  }
  return { type, params };
}

function collectSlides(body) {
  const lines = body.split(/\r?\n/);
  const slides = [];
  let cur = null;

  const push = () => {
    if (cur) slides.push(cur);
    cur = null;
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (line.startsWith("## ")) {
      push();
      cur = { title: line.replace(/^##\s+/, "").trim(), hint: null, content: [] };
      continue;
    }
    if (!cur) continue;

    const hint = parseHint(line);
    if (hint) {
      cur.hint = hint;
      continue;
    }
    cur.content.push(line);
  }
  push();
  return slides;
}

function parseBullets(contentLines) {
  const items = [];
  for (const l of contentLines) {
    const m = l.match(/^\s*-\s+(.*)\s*$/);
    if (m) items.push(m[1].trim());
  }
  return items.filter(Boolean);
}

function parseKeyValueList(items) {
  // "- key: value" -> { key, value }
  const out = [];
  for (const it of items) {
    const m = it.match(/^([^:]+)\s*:\s*(.+)$/);
    if (m) out.push({ k: m[1].trim(), v: m[2].trim() });
  }
  return out;
}

function parseMarkdownTable(contentLines) {
  // find consecutive table lines containing '|'
  const tableLines = contentLines.filter((l) => l.includes("|")).map((l) => l.trim());
  if (tableLines.length < 2) return null;

  // header
  const header = tableLines[0]
    .split("|")
    .map((s) => s.trim())
    .filter(Boolean);
  // separator line usually second
  const rows = [];
  for (let i = 2; i < tableLines.length; i++) {
    const cols = tableLines[i]
      .split("|")
      .map((s) => s.trim())
      .filter(Boolean);
    if (cols.length) rows.push(cols);
  }
  if (!header.length || !rows.length) return null;
  return { columns: header, rows };
}

export function mdToSpec(md, defaults = { defaultLayout: "LAYOUT_WIDE", defaultFont: "Aptos" }) {
  const { deck, body } = parseFrontmatter(md);
  const s = collectSlides(body);

  const layout = deck.layout || defaults.defaultLayout || "LAYOUT_WIDE";
  const fontFace = deck.themeFont || defaults.defaultFont || "Aptos";

  const spec = {
    layout,
    fileName: "output.pptx",
    theme: { fontFace },
    slides: []
  };

  for (const slide of s) {
    const hintType = slide.hint?.type || null;
    const params = slide.hint?.params || {};
    const bullets = parseBullets(slide.content);
    const table = parseMarkdownTable(slide.content);

    // Type decision (explicit first)
    let type = hintType;

    // Fallback inference when no explicit type
    if (!type) {
      if (table) type = "comparisonTable";
      else type = "bullets";
    }

    // Build per-type JSON
    if (type === "title") {
      // Expect "- title: X" "- subtitle: Y"
      const kv = parseKeyValueList(bullets);
      const title = kv.find((x) => x.k.toLowerCase() === "title")?.v || slide.title;
      const subtitle = kv.find((x) => x.k.toLowerCase() === "subtitle")?.v || "";
      spec.slides.push({ type: "title", title, subtitle });
      continue;
    }

    if (type === "imageHero") {
      spec.slides.push({
        type: "imageHero",
        title: slide.title,
        subtitle: params.subtitle || "",
        imagePath: params.imagePath || params.image || "",
        caption: params.caption || ""
      });
      continue;
    }

    if (type === "timeline") {
      // Use "- Label: Text" style
      const kv = parseKeyValueList(bullets);
      const items = kv.length
        ? kv.map((x) => ({ label: x.k, text: x.v }))
        : bullets.map((t, idx) => ({ label: `Step ${idx + 1}`, text: t }));
      spec.slides.push({ type: "timeline", title: slide.title, items });
      continue;
    }

    if (type === "kpiCards") {
      // "- 売上: 100" -> cards
      const kv = parseKeyValueList(bullets);
      const cards = kv.map((x) => ({ label: x.k, value: x.v }));
      spec.slides.push({ type: "kpiCards", title: slide.title, cards });
      continue;
    }

    if (type === "comparisonTable") {
      // prefer explicit columns param, else parse markdown table
      if (params.columns) {
        const columns = String(params.columns)
          .split(",")
          .map((x) => x.trim())
          .filter(Boolean);
        // rows from bullets as CSV: "- a, b, c"
        const rows = bullets
          .map((t) => t.split(",").map((x) => x.trim()))
          .filter((r) => r.length && r.some(Boolean));
        spec.slides.push({ type: "comparisonTable", title: slide.title, columns, rows });
      } else if (table) {
        spec.slides.push({ type: "comparisonTable", title: slide.title, columns: table.columns, rows: table.rows });
      } else {
        spec.slides.push({ type: "bullets", title: slide.title, items: bullets });
      }
      continue;
    }

    if (type === "chart") {
      // expects params.chartType and series encoded later (keep simple for now)
      // For now: if bullets are "Name: 1,2,3" then build series
      const chartType = (params.chartType || "bar").toLowerCase();
      const kv = parseKeyValueList(bullets);
      const series = kv.map((x) => {
        const nums = x.v.split(",").map((n) => Number(n.trim())).filter((n) => Number.isFinite(n));
        return { name: x.k, labels: nums.map((_, i) => `P${i + 1}`), values: nums };
      });
      spec.slides.push({ type: "chart", title: slide.title, chartType, series, options: { showLegend: true, showTitle: false } });
      continue;
    }

    if (type === "twoColumn") {
      // simple: first half bullets -> left, rest -> right
      const half = Math.ceil(bullets.length / 2);
      spec.slides.push({
        type: "twoColumn",
        title: slide.title,
        left: { heading: params.leftHeading || "要点", items: bullets.slice(0, half) },
        right: { heading: params.rightHeading || "補足", items: bullets.slice(half), text: params.rightText || "" }
      });
      continue;
    }

    // default bullets
    spec.slides.push({ type: "bullets", title: slide.title, items: bullets });
  }

  return spec;
}

// CLI usage for debug
if (process.argv[1]?.endsWith("md_to_spec.mjs")) {
  const args = process.argv.slice(2);
  const get = (k) => {
    const i = args.indexOf(k);
    return i >= 0 ? args[i + 1] : null;
  };
  const inPath = get("--in");
  const outPath = get("--out");
  if (!inPath || !outPath) {
    console.error("Usage: node md_to_spec.mjs --in slides.md --out spec.json");
    process.exit(1);
  }
  const fs = await import("node:fs");
  const md = fs.readFileSync(inPath, "utf-8");
  const spec = mdToSpec(md, { defaultLayout: "LAYOUT_WIDE", defaultFont: "Aptos" });
  fs.writeFileSync(outPath, JSON.stringify(spec, null, 2), "utf-8");
  console.log(`✅ wrote: ${outPath}`);
}
