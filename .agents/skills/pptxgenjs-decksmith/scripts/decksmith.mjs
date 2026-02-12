import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { mdToSpec } from "./md_to_spec.mjs";
import { generatePptx } from "./generate.mjs";

function parseArgs(argv) {
  const args = { in: null, out: "output.pptx", slides: null, spec: null };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--in") args.in = argv[++i];
    else if (a === "--out") args.out = argv[++i];
    else if (a === "--slides") args.slides = argv[++i];
    else if (a === "--spec") args.spec = argv[++i];
  }
  if (!args.in) {
    console.error("Usage: node decksmith.mjs --in <input.md|input.txt> [--out output.pptx] [--slides slides.md] [--spec spec.json]");
    process.exit(1);
  }
  return args;
}

// Very simple “Phase1”: input -> slides.md
// - If input already contains slide:type hints, we keep them.
// - If plain md/txt, we create bullets slides by H2 or paragraphs.
// You can improve this later with richer heuristics or LLM-based rewrite in the agent.
function toSlidesMd(raw, ext) {
  // If markdown already has slide:type hints, treat it as slides.md
  if (raw.includes("slide:type=")) return raw;

  // If markdown has H2 sections, make each H2 one slide (bullets)
  const lines = raw.split(/\r?\n/);
  const sections = [];
  let current = { title: null, body: [] };

  const pushCurrent = () => {
    if (current.title || current.body.some((l) => l.trim())) sections.push(current);
    current = { title: null, body: [] };
  };

  for (const line of lines) {
    if (line.startsWith("## ")) {
      pushCurrent();
      current.title = line.replace(/^##\s+/, "").trim();
    } else {
      current.body.push(line);
    }
  }
  pushCurrent();

  // If no H2, create a single slide
  if (sections.length === 0) {
    sections.push({ title: "Overview", body: lines });
  }

  // Convert section body to 3-6 bullets (naive: take first non-empty lines)
  const out = [];
  out.push("---");
  out.push("deck:");
  out.push("  layout: LAYOUT_WIDE");
  out.push("  themeFont: Aptos");
  out.push("---");
  out.push("");
  out.push("# Slide Deck");
  out.push("");

  // Title slide
  out.push("## タイトル");
  out.push("<!-- slide:type=title -->");
  out.push("- title: 自動生成スライド");
  out.push("- subtitle: Decksmith (MD→JSON→PPTX)");
  out.push("");

  for (const s of sections) {
    const title = s.title || "Section";
    const bullets = s.body
      .map((l) => l.trim())
      .filter((l) => l && !l.startsWith("#"))
      .slice(0, 6);

    out.push(`## ${title}`);
    out.push("<!-- slide:type=bullets -->");
    if (bullets.length === 0) {
      out.push("- （内容なし）");
    } else {
      for (const b of bullets) out.push(`- ${b}`);
    }
    out.push("");
  }

  return out.join("\n");
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const inputPath = path.resolve(process.cwd(), args.in);
  const outPptxPath = path.resolve(process.cwd(), args.out);

  const outDir = path.dirname(outPptxPath);
  const slidesPath = path.resolve(process.cwd(), args.slides || path.join(outDir, "slides.md"));
  const specPath = path.resolve(process.cwd(), args.spec || path.join(outDir, "spec.json"));

  const raw = fs.readFileSync(inputPath, "utf-8");
  const ext = path.extname(inputPath).toLowerCase();

  // Phase1: input -> slides.md
  const slidesMd = toSlidesMd(raw, ext);
  fs.writeFileSync(slidesPath, slidesMd, "utf-8");
  console.log(`✅ wrote slides.md: ${slidesPath}`);

  // Phase2: slides.md -> spec.json
  const spec = mdToSpec(slidesMd, { defaultLayout: "LAYOUT_WIDE", defaultFont: "Aptos" });
  fs.writeFileSync(specPath, JSON.stringify(spec, null, 2), "utf-8");
  console.log(`✅ wrote spec.json: ${specPath}`);

  // Phase3: spec.json -> pptx
  await generatePptx(spec, outPptxPath);
  console.log(`✅ wrote pptx: ${outPptxPath}`);
}

main().catch((e) => {
  console.error("❌ error:", e?.stack || e);
  process.exit(1);
});
