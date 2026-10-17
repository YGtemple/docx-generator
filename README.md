<div align="right">

**[English](README.md)** · [简体中文](README.zh-CN.md)

</div>

# docx-generator

Generate Word documents in code — a style system plus table, badge and divider components.

<sub>JavaScript · single dependency: `docx`</sub>

---

## Why not a template

The usual approach is to take a `.docx` template and replace placeholders. It has three problems:

1. **Changing style means opening Word.** Adjust a heading gap or an accent colour, and you are editing a binary file by hand.
2. **Dynamic content breaks placeholders.** Table rows that vary in count, paragraphs that appear only sometimes — the placeholder scheme collapses quickly.
3. **No diff.** A `.docx` is a zip. Change the template and there is nothing to review.

This takes the other route: **treat the document as a data structure**. The `docx` library's `Document` / `Paragraph` / `Table` API lets the document be declared in JavaScript.

What that buys:

- Styles live in one object; changing a colour updates the whole document
- Content is data-driven, so looping out table rows is natural
- The script is plain text — it goes into git, gets reviewed, diffs cleanly

## Design tokens

Colours are declared once, named by purpose rather than by hue:

```js
const ALGO = "F77F00";     // accent — algorithm / main line
const AIGC = "E63946";     // secondary accent — risk / emphasis
const NEU  = "3EC9A7";     // neutral — supplementary
const BG_CARD = "F5F3EE";  // card background
const FG = "1A1A24";       // primary text
const FG_DIM = "666666";   // secondary text
const BORDER_C = "DDDDDD"; // borders
```

Semantic names mean the theme can change ("algorithm" becomes "proposal") without renaming anything.

Chinese and Latin text need separate font slots:

```js
font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }
```

The `docx` font config has three slots. Set only one and Chinese falls back to a default face that looks visibly off next to Latin text at small sizes.

## Components

### Headings

```js
heading1(text)            // 36 half-points, primary
heading2(text, color)     // 28 half-points, accent by default
heading3(text, color)     // 24 half-points, primary text
```

Spacing is baked into the component (h1: `before: 360, after: 200`), so heading rhythm stays consistent across the whole document — the part hand-built Word documents rarely get right.

### Paragraphs

```js
para(text, { center, size, color, bold })   // single colour
richPara(runs, opts)                         // mixed runs in one paragraph
```

`richPara` is the important one. Different colours or weights inside one paragraph require multiple `TextRun`s in a single `Paragraph`. The component hides that.

### Info rows

```js
infoRow(key, val, valColor)
```

Produces a two-column "label | value" table where the label column is shaded:

```js
mkCell("Email", { w: 2200, shade: C.bgHead, c: C.dim, b: true, s: 18 })
mkCell("YOUR_EMAIL@example.com", { w: 6826, c: C.accent, b: true })
```

The widths (2200 / 6826) are not arbitrary — the A4 text area is about 9026 twips, split roughly 1:3 between label and value, which is the most balanced in practice.

### Badges and dividers

```js
chip(text, color)       // small shaded tag
sectionDivider()        // rule between sections
```

### One thing about borders

```js
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: BORDER_C };
```

`noBorder` must carry a `color` value even though it is unused. Without it, some Word versions draw a thin black line on cells that are supposed to have no border.

## Files

```
├── gen-docx.js       # v1 — complete, 30 KB, all components
├── gen-docx-v2.js    # v2 — refactored, 20 KB, denser abstractions
├── package.json
└── README.md
```

Both versions remain because v2 is not a strict replacement — v1's components are finer-grained and more customisable; v2 packs the common combinations tightly and writes faster.

## Usage

```bash
npm install docx
node gen-docx.js      # writes the document to the working directory
```

The generation logic lives in `main()` — edit it directly. One dependency: `docx`.

## Note

Contact details hard-coded in the original scripts were replaced with placeholders (`YOUR_EMAIL@example.com` / `YOUR_PHONE` / `YOUR_QQ`) before publishing. Change those three to reuse it.

## License

MIT
