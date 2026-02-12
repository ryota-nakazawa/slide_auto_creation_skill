---
name: pptxgenjs-decksmith
description: テキスト/Markdownから、指示付きスライドMarkdown(slides.md)→spec.json→pptxを生成する。レイアウト(type)を増やしながらリッチな資料を安定して量産するためのSkill。
---

# 目的
- ユーザーの入力（テキスト/Markdown）を「スライドに割り当て可能な構造」に変換し、PptxGenJSで .pptx を生成する。
- 変換の透明性とデバッグ性のために、中間生成物（slides.md / spec.json）を残す。

# 想定入力
- 素のテキスト（長文、箇条書き、メモ）
- Markdown（章立てされたレポート、Deep Research系など）

# 出力
- slides.md（スライド指示付きMarkdown）
- spec.json（pptx生成用JSON）
- output.pptx（最終生成物）

# 推奨ワークフロー（必ずこの順）
## Phase 0: 入力ファイルの確認
- 入力が .md か .txt かを確認
- 必要ならユーザーに「この入力は何枚くらいにしたいか」を聞く（既に指定があれば聞かない）

## Phase 1: 入力 → slides.md（指示付きMD）
- 章立て（H2）をスライド単位にし、1スライドあたり3〜6点に圧縮
- 以下のヒントをHTMLコメントで埋め込む（表示に影響しない）
  - <!-- slide:type=bullets -->
  - <!-- slide:type=twoColumn -->
  - <!-- slide:type=timeline -->
  - <!-- slide:type=comparisonTable columns=a,b,c -->
  - <!-- slide:type=kpiCards -->
  - <!-- slide:type=chart chartType=line -->
  - <!-- slide:type=imageHero imagePath=... -->

※ 指示がない場合は安全側（bullets）に倒す。

## Phase 2: slides.md → spec.json
- slide:type を優先して JSON の slides[].type にマップ
- tableや画像などの構造があれば type を推定する（ただし明示指定が最優先）

## Phase 3: spec.json → pptx
- scripts/generate.mjs で pptx を生成
- エラー時は必ず「slides.md」「spec.json」のどこが原因か切り分ける

# デバッグ指針
- 期待と違うレイアウトなら: slides.md に slide:type を明示追加する
- 期待と違う配置/見た目なら: generate.mjs の addXxx() を調整する
- 期待と違う分割なら: Phase1 の要約・章分割を調整する
