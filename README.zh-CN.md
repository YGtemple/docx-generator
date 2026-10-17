<div align="right">

[English](README.md) · **[简体中文](README.zh-CN.md)**

</div>

# docx-generator

用代码生成 Word 文档——一套样式系统，加上表格、徽章、分隔线组件。

<sub>JavaScript · 只有一个依赖：`docx`</sub>

---

## 为什么不用模板

常见做法是拿一个 `.docx` 模板替换占位符。这条路有三个麻烦：

1. **改样式要开 Word。** 调一下标题间距、换个强调色，都得手动编辑一个二进制文件。
2. **动态内容让占位符失控。** 表格行数不固定、某些段落只在特定情况下出现——占位符方案很快就撑不住。
3. **没有 diff。** `.docx` 是个压缩包。模板一改，什么都看不出来。

换个路子：**把文档当数据结构构造**。用 `docx` 库的 `Document` / `Paragraph` / `Table` API，在 JavaScript 里声明式地拼出一份文档。

换来的好处：

- 样式集中在一个对象里，改一个色值全局生效
- 内容由数据驱动，循环生成表格行是天然的
- 脚本是纯文本，能进 git、能 review、能 diff

## 设计令牌

配色在顶部集中声明，按用途命名而不是按颜色命名：

```js
const ALGO = "F77F00";     // 强调色——算法/主线
const AIGC = "E63946";     // 次强调——风险/重点
const NEU  = "3EC9A7";     // 中性——补充
const BG_CARD = "F5F3EE";  // 卡片底
const FG = "1A1A24";       // 主文字
const FG_DIM = "666666";   // 次要文字
const BORDER_C = "DDDDDD"; // 边框
```

用语义名而非 `RED`/`BLUE`，是为了让主题变化时（比如把「算法」换成「方案」）样式名不用跟着改。

中英文要分开指定字体：

```js
font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }
```

`docx` 的字体配置有 `ascii` / `hAnsi` / `eastAsia` 三个槽位。只设一个的话，中文会 fallback 到默认字体，和英文小字并排时明显不协调。

## 组件

### 标题

```js
heading1(text)            // 36 半磅，主色
heading2(text, color)     // 28 半磅，默认强调色
heading3(text, color)     // 24 半磅，主文字色
```

间距写死在组件里（一级 `before: 360, after: 200`），保证全文标题节奏一致——这是手写 Word 最难保证的部分。

### 段落

```js
para(text, { center, size, color, bold })   // 单色段落
richPara(runs, opts)                         // 段内多 run 混排
```

`richPara` 是关键。一段文字内不同部分用不同颜色或粗细，需要用多个 `TextRun` 组成一个 `Paragraph`。这个组件把那层封装掉了。

### 信息行

```js
infoRow(key, val, valColor)
```

生成「标签 | 值」两列表格，标签列带底色：

```js
mkCell("邮箱", { w: 2200, shade: C.bgHead, c: C.dim, b: true, s: 18 })
mkCell("YOUR_EMAIL@example.com", { w: 6826, c: C.accent, b: true })
```

宽度分配（2200 / 6826）不是随便给的——A4 正文区约 9026 twips，按约 1:3 分配标签与内容列，视觉上最平衡。

### 徽章与分隔线

```js
chip(text, color)       // 带底色的小标签
sectionDivider()        // 章节间横线
```

### 边框的一个坑

```js
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: BORDER_C };
```

`noBorder` 的 `color` 必须给个值（哪怕用不到），否则某些 Word 版本会在「无边框」单元格上渲染出黑色细线。

## 文件

```
├── gen-docx.js       # v1——完整版，30 KB，组件齐全
├── gen-docx-v2.js    # v2——重构版，20 KB，抽象更紧凑
├── package.json
└── README.md
```

两个版本并存，是因为 v2 不完全是 v1 的替代——v1 的组件更细碎、可定制性高；v2 把常用组合封装得更紧、写起来更快。

## 用法

```bash
npm install docx
node gen-docx.js      # 生成文档到当前目录
```

生成逻辑在 `main()` 里，直接改它即可。依赖只有 `docx` 一个。

## 说明

上传前已把脚本里硬编码的个人联系方式替换为占位符（`YOUR_EMAIL@example.com` / `YOUR_PHONE` / `YOUR_QQ`）。复用时改这几处即可。

## 许可证

MIT
