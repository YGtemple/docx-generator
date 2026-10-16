const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  Header, Footer, PageNumber, PageBreak, LevelFormat, TabStopType, TabStopPosition
} = require("docx");

// ═══ 设计令牌 ═══
// 集中定义配色，改一处全局生效。
// accent 用深琥珀而非亮橙——文档类排版需要克制、沉稳的强调色。
const C = {
  fg:     "1A1A2E",   // 主文字（近黑带蓝调，比纯黑柔和）
  dim:    "555566",   // 次要文字
  accent: "D4760A",   // 深琥珀——克制、沉稳
  line:   "D8D8D8",   // 分隔线与边框
  bgHead: "F6F4F0",   // 暖米色表头
  bgTag1: "FFF6EB",   // 浅琥珀标签底
  bgTag2: "FFEEF0",   // 浅玫红标签底
  bgTag3: "ECFDF5",   // 浅绿标签底
};

// ═══ 边框定义 ═══
// 注意 noLine 的 color 必须给值（哪怕用不到）——
// 某些 Word 版本会在「无边框」单元格上渲染出黑色细线。
const thin = { style: BorderStyle.SINGLE, size: 1, color: C.line };
const thinB = { top: thin, bottom: thin, left: thin, right: thin };
const noLine = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noB = { top: noLine, bottom: noLine, left: noLine, right: noLine };
const accentB = { style: BorderStyle.SINGLE, size: 6, color: C.accent, space: 1 };

/** 中英文字体三槽位（只设一个中文会 fallback，与英文并排不协调） */
const FONT = { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" };

/**
 * 表格单元格工厂。
 * text 支持两种形态：
 *   - 字符串          → 单 run
 *   - [{ t, c, b, s }] → 多 run 混排（段内变色/变粗）
 */
function mkCell(text, opts = {}) {
  const runs = Array.isArray(text)
    ? text.map(t => new TextRun({
        text: t.t,
        size: t.s || opts.s || 20,
        color: t.c || opts.c || C.fg,
        bold: t.b || opts.b,
        font: FONT
      }))
    : [new TextRun({
        text,
        size: opts.s || 20,
        color: opts.c || C.fg,
        bold: opts.b,
        font: FONT
      })];
  return new TableCell({
    borders: opts.borders || thinB,
    width: opts.w ? { size: opts.w, type: WidthType.DXA } : undefined,
    shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: opts.mt || 70, bottom: opts.mb || 70, left: opts.ml || 130, right: opts.mr || 130 },
    verticalAlign: opts.va || "center",
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: runs
    })]
  });
}

/** 段落工厂。同样支持字符串或多 run 数组。 */
function p(text, opts = {}) {
  const runs = Array.isArray(text)
    ? text.map(t => new TextRun({
        text: t.t, size: t.s || 22, color: t.c || C.dim, bold: t.b, font: FONT
      }))
    : [new TextRun({
        text, size: opts.s || 22, color: opts.c || C.dim, bold: opts.b, font: FONT
      })];
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: opts.before || 40, after: opts.after || 40 },
    children: runs
  });
}

/** 章节分隔线（用段落下边框实现，比插入图形更可控） */
function divider() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: C.line, space: 8 } },
    children: []
  });
}

/**
 * 带编号的章节标题。num 与 text 分色——
 * 编号用 accent 色，标题用主文字色，视觉层次清楚。
 */
function sectionTitle(num, text) {
  return new Paragraph({
    spacing: { before: 280, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.accent, space: 6 } },
    children: [
      new TextRun({ text: `${num}`, size: 30, bold: true, color: C.accent, font: FONT }),
      new TextRun({ text: `  ${text}`, size: 28, bold: true, color: C.fg, font: FONT }),
    ]
  });
}

/** 标签徽章：带底色的小标签，用于分类标记 */
function chip(text, bg = C.bgTag1, fg = C.accent) {
  return new TableCell({
    borders: noB,
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 40, bottom: 40, left: 100, right: 100 },
    width: { size: 1400, type: WidthType.DXA },
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, size: 17, bold: true, color: fg, font: FONT })]
    })]
  });
}

/**
 * 「标签 | 值」信息行。
 *
 * 列宽 2200 / 6826：A4 正文区约 9026 twips，
 * 按约 1:3 分配，视觉上最平衡。
 */
function infoRow(key, val, valColor) {
  return new TableRow({
    cantSplit: true,
    children: [
      mkCell(key, { w: 2200, shade: C.bgHead, c: C.dim, b: true, s: 18 }),
      mkCell(val, { w: 6826, c: valColor || C.fg, b: true, s: 20 })
    ]
  });
}

/** 三栏要点表格（用于并列对比） */
function tripleTable(rows) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    borders: thinB,
    rows: [
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: [
          mkCell("维度", { w: 2200, shade: C.bgHead, c: C.dim, b: true, s: 18 }),
          mkCell("说明", { w: 4413, shade: C.bgHead, c: C.dim, b: true, s: 18 }),
          mkCell("备注", { w: 2413, shade: C.bgHead, c: C.dim, b: true, s: 18 }),
        ]
      }),
      ...rows.map(r => new TableRow({
        cantSplit: true,
        children: [
          mkCell(r[0], { w: 2200, c: C.fg, b: true }),
          mkCell(r[1], { w: 4413, c: C.dim }),
          mkCell(r[2] || "", { w: 2413, c: C.dim, s: 18 }),
        ]
      }))
    ]
  });
}

/** 页脚：居中页码 */
function buildFooter() {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        children: ["第 ", PageNumber.CURRENT, " 页 / 共 ", PageNumber.TOTAL_PAGES, " 页"],
        size: 16, color: C.dim, font: FONT
      })]
    })]
  });
}

/**
 * 主流程：按顺序拼装文档区块。
 * 改成生成别的内容，只需改这里。
 */
async function main() {
  const children = [];

  // ── 标题区 ──
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text: "文档标题", size: 44, bold: true, color: C.fg, font: FONT })]
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 240 },
    children: [new TextRun({ text: "副标题 / 说明行", size: 20, color: C.dim, font: FONT })]
  }));
  children.push(divider());

  // ── 基本信息表 ──
  children.push(sectionTitle("01", "基本信息"));
  children.push(new Table({
    width: { size: 9026, type: WidthType.DXA },
    borders: noB,
    rows: [
      // 上传前已脱敏：这里原本是真实邮箱/手机号
      infoRow("邮箱", "YOUR_EMAIL@example.com", C.accent),
      infoRow("电话", "YOUR_PHONE（添加时请注明来意）"),
    ]
  }));

  // ── 要点列表 ──
  children.push(sectionTitle("02", "核心要点"));
  const points = [
    "第一条要点，用 para 组件生成",
    "第二条要点",
    "第三条要点",
  ];
  for (const t of points) children.push(p("· " + t));

  // ── 对比表 ──
  children.push(sectionTitle("03", "对比说明"));
  children.push(tripleTable([
    ["维度 A", "这里是维度的说明文字，可用多 run 混排来强调关键短语", "备注"],
    ["维度 B", "同上", ""],
  ]));
  children.push(divider());

  // ── 组装 ──
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: FONT, size: 22, color: C.fg } }
      }
    },
    sections: [{
      properties: {
        page: {
          margin: { top: 1000, right: 1100, bottom: 1000, left: 1100 }
        }
      },
      footers: { default: buildFooter() },
      children
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  const outPath = "output.docx";
  fs.writeFileSync(outPath, buffer);
  console.log(`[OK] 已生成 ${outPath} (${(buffer.length / 1024).toFixed(1)} KB)`);
}

main().catch(err => {
  console.error("[FAIL]", err);
  process.exit(1);
});
