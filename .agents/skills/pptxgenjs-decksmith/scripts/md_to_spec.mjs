function parseFrontmatter(md) {
  const m = md.match(/^---\s*\n([\s\S]*?)\n---\s*\n/);
  if (!m) return { deck: {}, body: md };
  const fm = m[1];
  const body = md.slice(m[0].length);

  const deck = {};
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
  // <!-- slide:template=sidebarCards iconChar=... -->
  const m = line.match(/<!--\s*slide:([a-zA-Z]+)=([a-zA-Z0-9_]+)([\s\S]*?)-->/);
  if (!m) return null;
  const kind = m[1].trim(); // type or template
  const value = m[2].trim();
  const rest = (m[3] || "").trim();

  const params = {};
  const re = /([a-zA-Z0-9_]+)\s*=\s*("([^"]+)"|[^\s]+)/g;
  let mm;
  while ((mm = re.exec(rest))) {
    const key = mm[1];
    const raw = mm[3] ?? mm[2];
    params[key] = String(raw).replace(/^"|"$/g, "");
  }
  return { kind, value, params };
}

function collectSlides(body) {
  const lines = body.split(/\r?\n/);
  const slides = [];
  let cur = null;

  const push = () => {
    if (cur) slides.push(cur);
    cur = null;
  };

  for (const line of lines) {
    if (line.startsWith("## ")) {
      push();
      cur = { title: line.replace(/^##\s+/, "").trim(), hints: {}, params: {}, raw: [] };
      continue;
    }
    if (!cur) continue;

    const hint = parseHint(line);
    if (hint) {
      cur.hints[hint.kind] = hint.value;
      cur.params = { ...cur.params, ...hint.params };
      continue;
    }

    cur.raw.push(line);
  }
  push();
  return slides;
}

/**
 * Parse subsections:
 * ### Heading
 * - bullet
 */
function parseSections(rawLines) {
  const blocks = [];
  let cur = null;

  const push = () => {
    if (cur) blocks.push(cur);
    cur = null;
  };

  for (const line of rawLines) {
    if (line.startsWith("### ")) {
      push();
      cur = { heading: line.replace(/^###\s+/, "").trim(), bullets: [] };
      continue;
    }
    if (!cur) continue;
    const m = line.match(/^\s*-\s+(.*)\s*$/);
    if (m) cur.bullets.push(m[1].trim());
  }
  push();
  return blocks;
}

/**
 * Parse table blocks inside a section (optional):
 * We support markdown tables like:
 * | A | B | C |
 * |---|---|---|
 * | 1 | 2 | 3 |
 */
function parseMarkdownTable(rawLines) {
  const rows = [];
  for (const l of rawLines) {
    if (!l.trim().startsWith("|")) continue;
    const cols = l.trim().replace(/^\||\|$/g, "").split("|").map((c) => c.trim());
    rows.push(cols);
  }
  if (rows.length < 2) return null;

  // Remove separator row if it looks like ---|---|
  const header = rows[0];
  const body = rows.slice(1).filter((r) => !r.every((c) => /^-+$/.test(c.replace(/:/g, "").trim())));
  return { header, rows: body };
}

export function mdToSpec(md, defaults = { defaultLayout: "LAYOUT_WIDE", defaultFont: "Aptos" }) {
  const { deck, body } = parseFrontmatter(md);
  const slidesMd = collectSlides(body);

  const spec = {
    layout: deck.layout || defaults.defaultLayout || "LAYOUT_WIDE",
    fileName: "output.pptx",
    theme: { fontFace: deck.themeFont || defaults.defaultFont || "Aptos" },
    slides: []
  };

  for (const s of slidesMd) {
    const template = s.hints.template || null;
    const type = s.hints.type || null;
    const blocks = parseSections(s.raw);

    // Title slide shortcut
    if ((type || template) === "title") {
      const lines = s.raw.filter((l) => l.trim().startsWith("-"));
      const kv = {};
      for (const l of lines) {
        const mm = l.replace(/^\s*-\s*/, "").match(/^([^:]+)\s*:\s*(.+)$/);
        if (mm) kv[mm[1].trim().toLowerCase()] = mm[2].trim();
      }
      spec.slides.push({ template: "title", title: kv.title || s.title, subtitle: kv.subtitle || "" });
      continue;
    }

    // Template-first
    if (template) {
      spec.slides.push({
        template,
        title: s.title,
        params: s.params || {},
        blocks
      });
      continue;
    }

    // Type-based slides (kept for backward compat; generate will map these to templates)
    // Enhance spec for known types
    if (type === "twoColumn") {
      // Expect blocks: ### 左 / ### 右 (or any 2 blocks)
      const left = blocks[0] || { heading: "左", bullets: [] };
      const right = blocks[1] || { heading: "右", bullets: [] };
      spec.slides.push({
        type,
        title: s.title,
        params: s.params || {},
        left: { heading: left.heading, bullets: left.bullets },
        right: { heading: right.heading, bullets: right.bullets }
      });
      continue;
    }

    if (type === "timeline") {
      // Each block is a step
      const steps = blocks.map((b) => ({ title: b.heading, body: b.bullets }));
      spec.slides.push({
        type,
        title: s.title,
        params: s.params || {},
        steps
      });
      continue;
    }

    if (type === "comparisonTable") {
      // First try parse markdown table from raw
      const table = parseMarkdownTable(s.raw);
      if (table) {
        spec.slides.push({
          type,
          title: s.title,
          params: s.params || {},
          headers: table.header,
          rows: table.rows
        });
        continue;
      }
      // Fallback: use blocks as rows [heading + bullets...]
      const headers = ["項目", "内容"];
      const rows = blocks.map((b) => [b.heading, (b.bullets || []).join(" / ")]);
      spec.slides.push({ type, title: s.title, params: s.params || {}, headers, rows });
      continue;
    }

    if (type) {
      spec.slides.push({ type, title: s.title, params: s.params || {}, blocks });
      continue;
    }

    // Default bullets if nothing is specified
    const bullets = blocks.flatMap((b) => b.bullets);
    spec.slides.push({ type: "bullets", title: s.title, items: bullets });
  }

  return spec;
}

// CLI debug
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
