# pptxgenjs-decksmith

## Install
npm install

## Run (recommended)
node decksmith.mjs --in ../../../../input.md --out ../../../../output.pptx

Outputs:
- slides.md (same directory as output by default)
- spec.json (same directory as output by default)

## Run each step (debug)
node md_to_spec.mjs --in slides.md --out spec.json
node generate.mjs spec.json output.pptx
