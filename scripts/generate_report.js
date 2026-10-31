// 幼儿园集团 — 2026年10月园务设备信息化管理月度工作报告
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  TableLayoutType, Footer, PageNumber, SectionType, NumberFormat,
} = require("docx");
const fs = require("fs");

// ── WM-1 Warm Teal（教育培训）──
const P = {
  bg: "F4F1E9", accent: "FF6A3B",
  cover: { titleColor: "15857A", subtitleColor: "606060", metaColor: "707070", footerColor: "A0A0A0" },
  table: { headerBg: "15857A", headerText: "FFFFFF", innerLine: "D5D0C8", surface: "F0EDE5" },
  primary: "15857A",
};

const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const allNoBorders = { ...noBorders, insideHorizontal: NB, insideVertical: NB };

// ── calcTitleLayout / splitTitleLines / calcCoverSpacing（design-system 标准实现）──
function splitTitleLines(title, charsPerLine) {
  if (title.length <= charsPerLine) return [title];
  const breakAfter = new Set([..."\uff0c\u3002\u3001\uff1b\uff1a\uff01\uff1f", ..."\u7684\u4e0e\u548c\u53ca\u4e4b\u5728\u4e8e\u4e3a", ..."-_\u2014\u2013\u00b7/", ..." \t"]);
  const lines = []; let remaining = title;
  while (remaining.length > charsPerLine) {
    let breakAt = -1;
    for (let i = charsPerLine; i >= Math.floor(charsPerLine * 0.6); i--) {
      if (i < remaining.length && breakAfter.has(remaining[i - 1])) { breakAt = i; break; }
    }
    if (breakAt === -1) {
      const limit = Math.min(remaining.length, Math.ceil(charsPerLine * 1.3));
      for (let i = charsPerLine + 1; i < limit; i++) {
        if (breakAfter.has(remaining[i - 1])) { breakAt = i; break; }
      }
    }
    if (breakAt === -1) {
      breakAt = charsPerLine;
      const prev = remaining[breakAt - 1], next = remaining[breakAt];
      if (prev && next && !breakAfter.has(prev) && !breakAfter.has(next) &&
          /[\u4e00-\u9fff]/.test(prev) && /[\u4e00-\u9fff]/.test(next)) breakAt -= 1;
    }
    lines.push(remaining.slice(0, breakAt).trim());
    remaining = remaining.slice(breakAt).trim();
  }
  if (remaining) lines.push(remaining);
  if (lines.length > 1 && lines[lines.length - 1].length <= 2) {
    const last = lines.pop(); lines[lines.length - 1] += last;
  }
  return lines;
}
function calcTitleLayout(title, maxWidthTwips, preferredPt = 40, minPt = 24) {
  const charsPerLine = (pt) => Math.floor(maxWidthTwips / (pt * 20));
  let titlePt = preferredPt, lines;
  while (titlePt >= minPt) {
    const cpl = charsPerLine(titlePt);
    if (cpl < 2) { titlePt -= 2; continue; }
    lines = splitTitleLines(title, cpl);
    if (lines.length <= 3) break;
    titlePt -= 2;
  }
  if (!lines || lines.length > 3) { lines = splitTitleLines(title, charsPerLine(minPt)); titlePt = minPt; }
  return { titlePt, titleLines: lines };
}
function calcCoverSpacing(p) {
  const { titleLineCount = 1, titlePt = 36, hasSubtitle = false, hasEnglishLabel = false,
          metaLineCount = 0, fixedHeight = 800, pageHeight = 16838 } = p;
  const SAFETY = 1200;
  const usable = pageHeight - SAFETY;
  const contentH = titleLineCount * (titlePt * 23 + 200) + (hasSubtitle ? 12 * 23 + 600 : 0) +
    (hasEnglishLabel ? 9 * 23 + 600 : 0) + metaLineCount * (10 * 23 + 100) + fixedHeight + 3 * 300;
  const safeRemaining = Math.max(usable - contentH, 400);
  const FOOTER_MIN = 800;
  const rawTop = Math.floor(safeRemaining * 0.45), rawBottom = Math.floor(safeRemaining * 0.45);
  const bottomSpacing = Math.max(rawBottom, FOOTER_MIN);
  const topSpacing = Math.max(rawTop - Math.max(0, FOOTER_MIN - rawBottom), 400);
  return { topSpacing, bottomSpacing };
}

// ── Recipe R1 封面 ──
function buildCoverR1(config) {
  const C = config.palette;
  const padL = 1200, padR = 800;
  const availableWidth = 11906 - padL - padR - 300;
  const { titlePt, titleLines } = calcTitleLayout(config.title, availableWidth, 40, 24);
  const titleSize = titlePt * 2;
  const spacing = calcCoverSpacing({
    titleLineCount: titleLines.length, titlePt,
    hasSubtitle: !!config.subtitle, hasEnglishLabel: !!config.englishLabel,
    metaLineCount: (config.metaLines || []).length, fixedHeight: 400,
  });
  const accentLeft = { style: BorderStyle.SINGLE, size: 8, color: C.accent, space: 12 };
  const children = [];
  children.push(new Paragraph({ spacing: { before: spacing.topSpacing } }));
  if (config.englishLabel) {
    children.push(new Paragraph({
      indent: { left: padL, right: padR }, spacing: { after: 500 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.accent, space: 8 } },
      children: [new TextRun({ text: config.englishLabel.split("").join("  "), size: 18,
        color: C.accent, font: { ascii: "Calibri", eastAsia: "SimHei" }, characterSpacing: 40 })],
    }));
  }
  for (let i = 0; i < titleLines.length; i++) {
    children.push(new Paragraph({
      indent: { left: padL },
      spacing: { after: i < titleLines.length - 1 ? 100 : 300, line: Math.ceil(titlePt * 23), lineRule: "atLeast" },
      children: [new TextRun({ text: titleLines[i], size: titleSize, bold: true,
        color: C.cover.titleColor, font: { eastAsia: "SimHei", ascii: "Arial" } })],
    }));
  }
  if (config.subtitle) {
    children.push(new Paragraph({
      indent: { left: padL }, spacing: { after: 800 },
      children: [new TextRun({ text: config.subtitle, size: 24, color: C.cover.subtitleColor,
        font: { eastAsia: "Microsoft YaHei", ascii: "Arial" } })],
    }));
  }
  for (const line of (config.metaLines || [])) {
    children.push(new Paragraph({
      indent: { left: padL + 200 }, spacing: { after: 80 },
      border: { left: accentLeft },
      children: [new TextRun({ text: line, size: 24, color: C.cover.metaColor,
        font: { eastAsia: "Microsoft YaHei", ascii: "Arial" } })],
    }));
  }
  children.push(new Paragraph({ spacing: { before: spacing.bottomSpacing } }));
  children.push(new Paragraph({
    indent: { left: padL, right: padR },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: C.accent, space: 8 } },
    spacing: { before: 200 },
    children: [
      new TextRun({ text: config.footerLeft || "", size: 16, color: C.cover.footerColor, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
      new TextRun({ text: "                                        " }),
      new TextRun({ text: config.footerRight || "", size: 16, color: C.cover.footerColor, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
    ],
  }));
  return [new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: allNoBorders,
    rows: [new TableRow({
      height: { value: 16838, rule: "exact" },
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: config.palette.bg }, borders: noBorders,
        children,
      })],
    })],
  })];
}

// ── 正文构件（Profile A 正式风格）──
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160, line: 312 },
    children: [new TextRun({ text, bold: true, size: 32, color: P.primary, font: { ascii: "Times New Roman", eastAsia: "SimHei" } })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120, line: 312 },
    children: [new TextRun({ text, bold: true, size: 28, color: "000000", font: { ascii: "Times New Roman", eastAsia: "SimHei" } })],
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 480 },
    spacing: { line: 312, after: 80 },
    children: [new TextRun({ text, size: 24, color: "000000", font: { ascii: "Times New Roman", eastAsia: "SimSun" }, bold: !!opts.bold })],
  });
}
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 60, after: 240, line: 312 },
    children: [new TextRun({ text, size: 21, color: "606060", font: { ascii: "Times New Roman", eastAsia: "SimSun" } })],
  });
}
function tableTitle(text) {
  return new Paragraph({
    keepNext: true, spacing: { before: 200, after: 100, line: 312 },
    children: [new TextRun({ text, bold: true, size: 21, color: "000000", font: { ascii: "Times New Roman", eastAsia: "SimHei" } })],
  });
}
function cellPara(text, { bold = false, header = false, right = false } = {}) {
  return new Paragraph({
    alignment: right ? AlignmentType.RIGHT : (header ? AlignmentType.CENTER : AlignmentType.LEFT),
    spacing: { line: 276 },
    children: [new TextRun({ text: String(text), size: 21, bold: bold || header,
      color: header ? P.table.headerText : "000000",
      font: { ascii: "Times New Roman", eastAsia: header ? "SimHei" : "SimSun" } })],
  });
}
function dataTable(headers, rows, widths, rightCols = []) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: P.table.headerBg },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: P.table.headerBg },
      left: NB, right: NB,
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: P.table.innerLine },
      insideVertical: NB,
    },
    rows: [
      new TableRow({
        tableHeader: true, cantSplit: true,
        children: headers.map((h, i) => new TableCell({
          width: { size: widths[i], type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: P.table.headerBg },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [cellPara(h, { header: true })],
        })),
      }),
      ...rows.map((r, ri) => new TableRow({
        cantSplit: true,
        children: r.map((v, i) => new TableCell({
          width: { size: widths[i], type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: ri % 2 === 0 ? "FFFFFF" : P.table.surface },
          margins: { top: 70, bottom: 70, left: 120, right: 120 },
          children: [cellPara(v, { right: rightCols.includes(i), bold: String(r[0]).includes("\u5408\u8ba1") })],
        })),
      })),
    ],
  });
}

// ── 正文内容 ──
const bodyChildren = [
  h1("\u4e00\u3001\u672c\u6708\u5de5\u4f5c\u6982\u8ff0"),
  body("\u672c\u62a5\u544a\u671f\u4e3a2026\u5e7410\u6708\u3002\u672c\u6708\u96c6\u56e2\u4e09\u6240\u5e7c\u513f\u56ed\uff08\u9633\u5149\u56ed\u3001\u5f69\u8679\u56ed\u3001\u661f\u8fb0\u56ed\uff09\u7ee7\u7eed\u6309\u7167\u300a\u56ed\u52a1\u8bbe\u5907\u7ba1\u7406\u7cfb\u7edf\u300b\u5f00\u5c55\u8bbe\u5907\u53f0\u8d26\u7ef4\u62a4\u4e0e\u7ef4\u4fee\u7ba1\u7406\u5de5\u4f5c\uff0c\u5168\u90e8\u6570\u636e\u5df2\u5f55\u5165\u7edf\u4e00\u7684 Excel \u6570\u636e\u5e93\u5e76\u7531\u516c\u5f0f\u81ea\u52a8\u6c47\u603b\u3002\u622a\u81f3\u672c\u6708\u672b\uff0c\u96c6\u56e2\u5728\u518c\u8bbe\u5907\u5171 32 \u53f0\uff0c\u5176\u4e2d\u5728\u7528 23 \u53f0\uff08\u5360 71.9%\uff09\uff0c\u7ef4\u4fee\u4e2d 4 \u53f0\uff0c\u95f2\u7f6e 3 \u53f0\uff0c\u5f85\u62a5\u5e9f 2 \u53f0\u3002"),
  body("6\u201410\u6708\u7d2f\u8ba1\u53d7\u7406\u62a5\u4fee\u5355 20 \u5355\uff0c\u5df2\u5b8c\u6210 15 \u5355\uff0c\u7ef4\u4fee\u5b8c\u6210\u7387 75.0%\uff0c\u7d2f\u8ba1\u7ef4\u4fee\u8d39\u7528 6,740.00 \u5143\u3002\u672c\u6708\u65b0\u589e\u62a5\u4fee 4 \u5355\uff0c\u5df2\u5b8c\u6210 1 \u5355\uff0c\u5176\u4f59 3 \u5355\u6b63\u5728\u7ef4\u4fee\u6216\u5f85\u6d3e\u5355\uff0c\u9884\u8ba1\u5c06\u4e8e11\u6708\u4e0a\u65ec\u5168\u90e8\u529e\u7ed3\u3002"),

  h1("\u4e8c\u3001\u8bbe\u5907\u8d44\u4ea7\u603b\u4f53\u60c5\u51b5"),
  tableTitle("\u8868 1  \u5404\u56ed\u533a\u8bbe\u5907\u6982\u51b5\uff08\u622a\u81f3 2026 \u5e74 10 \u6708\uff09"),
  dataTable(
    ["\u56ed\u533a", "\u8bbe\u5907\u603b\u6570", "\u5728\u7528", "\u7ef4\u4fee\u4e2d", "\u95f2\u7f6e", "\u5f85\u62a5\u5e9f", "\u7ef4\u4fee\u4e2d\u5360\u6bd4"],
    [
      ["\u9633\u5149\u56ed", "12", "9", "2", "1", "0", "16.7%"],
      ["\u5f69\u8679\u56ed", "11", "8", "1", "1", "1", "9.1%"],
      ["\u661f\u8fb0\u56ed", "9", "6", "1", "1", "1", "11.1%"],
      ["\u5408\u8ba1", "32", "23", "4", "3", "2", "12.5%"],
    ],
    [16, 14, 12, 13, 12, 13, 20], [1, 2, 3, 4, 5, 6]
  ),
  caption("\u6570\u636e\u6765\u6e90\uff1a\u300a\u56ed\u52a1\u8bbe\u5907\u7ba1\u7406\u7cfb\u7edf.xlsx\u300b\u201c\u8bbe\u5907\u53f0\u8d26\u201d\u201c\u6708\u5ea6\u7edf\u8ba1\u201d\u5de5\u4f5c\u8868"),
  body("\u4ece\u56ed\u533a\u5206\u5e03\u770b\uff0c\u9633\u5149\u56ed\u8bbe\u5907\u5b58\u91cf\u6700\u5927\uff0812 \u53f0\uff09\uff0c\u4e14\u7ef4\u4fee\u4e2d\u8bbe\u5907\u4e5f\u76f8\u5bf9\u8f83\u591a\uff082 \u53f0\uff09\uff1b\u4ece\u4f7f\u7528\u72b6\u6001\u770b\uff0c\u5168\u96c6\u56e2\u8bbe\u5907\u5728\u7528\u7387 71.9%\uff0c\u6574\u4f53\u8fd0\u884c\u826f\u597d\u3002\u5404\u56ed\u5747\u5b58\u5728 1 \u53f0\u95f2\u7f6e\u8bbe\u5907\uff0c\u5efa\u8bae\u7ed3\u5408\u4e0b\u5b63\u5ea6\u73ed\u989d\u8c03\u6574\u7edf\u4e00\u8c03\u62e8\u4f7f\u7528\uff1b\u5f69\u8679\u56ed\u591a\u5a92\u4f53\u6295\u5f71\u4eea\u4e0e\u661f\u8fb0\u56ed\u9488\u5f0f\u6253\u5370\u673a\u5df2\u8fbe\u62a5\u5e9f\u6761\u4ef6\uff0c\u5df2\u5217\u5165\u672c\u5b66\u671f\u8d44\u4ea7\u5904\u7f6e\u7533\u62a5\u8ba1\u5212\u3002"),

  h1("\u4e09\u3001\u7ef4\u4fee\u5de5\u4f5c\u4e0e\u8d39\u7528\u5206\u6790"),
  h2("\uff08\u4e00\uff09\u6708\u5ea6\u62a5\u4fee\u4e0e\u5b8c\u6210\u8d8b\u52bf"),
  tableTitle("\u8868 2  2026 \u5e74 6\u201410 \u6708\u62a5\u4fee\u4e0e\u7ef4\u4fee\u5b8c\u6210\u60c5\u51b5"),
  dataTable(
    ["\u6708\u4efd", "\u62a5\u4fee\u6570\u91cf", "\u5df2\u5b8c\u6210", "\u5b8c\u6210\u7387", "\u7ef4\u4fee\u8d39\u7528(\u5143)"],
    [
      ["2026-06", "3", "3", "100.0%", "890.00"],
      ["2026-07", "4", "4", "100.0%", "2,620.00"],
      ["2026-08", "4", "4", "100.0%", "1,320.00"],
      ["2026-09", "5", "3", "60.0%", "1,330.00"],
      ["2026-10", "4", "1", "25.0%", "580.00"],
      ["\u5408\u8ba1", "20", "15", "75.0%", "6,740.00"],
    ],
    [20, 20, 18, 20, 22], [1, 2, 3, 4]
  ),
  caption("\u6570\u636e\u6765\u6e90\uff1a\u300a\u56ed\u52a1\u8bbe\u5907\u7ba1\u7406\u7cfb\u7edf.xlsx\u300b\u201c\u7ef4\u4fee\u8bb0\u5f55\u201d\u5de5\u4f5c\u8868\uff08\u516c\u5f0f\u5b9e\u65f6\u7edf\u8ba1\uff09"),
  body("6\u20148\u6708\u62a5\u4fee\u5355\u5747\u5f53\u6708\u529e\u7ed3\uff0c\u5b8c\u6210\u7387\u4fdd\u6301 100%\uff1b9 \u6708\u8d77\u56e0\u62a5\u4fee\u91cf\u4e0a\u5347\u53ca\u4e24\u53f0\u8bbe\u5907\u9700\u5916\u5305\u7ef4\u4fee\uff0c\u5b8c\u6210\u7387\u9010\u6708\u56de\u843d\uff0c10 \u6708\u5b8c\u6210\u7387\u4e3a 25.0%\u3002\u9700\u8981\u8bf4\u660e\u7684\u662f\uff0c10 \u6708 3 \u5355\u672a\u5b8c\u6210\u7ef4\u4fee\u4e2d\u6709 2 \u5355\u4e3a\u6708\u5e95\u65b0\u62a5\uff0c\u5c5e\u6b63\u5e38\u5728\u529e\u5468\u671f\u5185\uff1b\u4f46\u9633\u5149\u56ed\u53f0\u5f0f\u7535\u8111\uff08SB-004\uff09\u81ea 9 \u6708\u5e95\u62a5\u4fee\u81f3\u4eca\u4ecd\u5f85\u6d3e\u5355\uff0c\u5df2\u8d85\u51fa\u5e38\u89c4\u5468\u671f\uff0c\u9700\u91cd\u70b9\u8ddf\u8fdb\u3002"),
  h2("\uff08\u4e8c\uff09\u6545\u969c\u7c7b\u578b\u5206\u5e03"),
  tableTitle("\u8868 3  \u6545\u969c\u7c7b\u578b\u5206\u5e03\u4e0e\u8d39\u7528\u6784\u6210\uff082026 \u5e74 6\u201410 \u6708\uff09"),
  dataTable(
    ["\u6545\u969c\u7c7b\u578b", "\u62a5\u4fee\u6570\u91cf", "\u7ef4\u4fee\u8d39\u7528(\u5143)", "\u8d39\u7528\u5360\u6bd4"],
    [
      ["\u673a\u68b0\u635f\u574f", "3", "2,480.00", "36.8%"],
      ["\u5176\u4ed6\uff08\u7ba1\u8def\u6f0f\u6c34\u3001\u5149\u5b66\u90e8\u4ef6\u7b49\uff09", "4", "1,550.00", "23.0%"],
      ["\u5236\u51b7\u5f02\u5e38", "3", "1,200.00", "17.8%"],
      ["\u7535\u8def\u6545\u969c", "3", "500.00", "7.4%"],
      ["\u89e6\u63a7\u5931\u7075", "2", "460.00", "6.8%"],
      ["\u8017\u6750\u8001\u5316", "2", "330.00", "4.9%"],
      ["\u8f6f\u4ef6\u6545\u969c", "3", "220.00", "3.3%"],
      ["\u5408\u8ba1", "20", "6,740.00", "100.0%"],
    ],
    [40, 18, 22, 20], [1, 2, 3]
  ),
  body("\u673a\u68b0\u635f\u574f\u4e3b\u8981\u96c6\u4e2d\u4e8e\u6237\u5916\u5927\u578b\u8fd0\u52a8\u8bbe\u65bd\uff0c\u5355\u4ef6\u7ef4\u4fee\u8d39\u7528\u9ad8\uff0c\u662f\u8d39\u7528\u9996\u8981\u6765\u6e90\uff08\u5360 36.8%\uff09\uff1b\u5236\u51b7\u5f02\u5e38\u96c6\u4e2d\u4e8e\u4f7f\u7528\u5e74\u9650\u8d85\u8fc7 5 \u5e74\u7684\u7a7a\u8c03\u8bbe\u5907\uff0c\u5efa\u8bae\u5165\u51ac\u524d\u5b8c\u6210\u4e00\u6b21\u6027\u96c6\u4e2d\u4fdd\u517b\uff0c\u964d\u4f4e\u53d6\u6696\u5b63\u6545\u969c\u7387\u3002"),

  h1("\u56db\u3001\u4e3b\u8981\u95ee\u9898\u4e0e\u98ce\u9669\u63d0\u793a"),
  body("\u4e00\u662f\u7ef4\u4fee\u65f6\u6548\u6709\u6ce2\u52a8\u30029\u201410\u6708\u5b8c\u6210\u7387\u4e0b\u964d\uff0c\u53cd\u6620\u7ef4\u4fee\u8d44\u6e90\uff08\u5c24\u5176\u5916\u5305\u7ef4\u4fee\u5546\u54cd\u5e94\u901f\u5ea6\uff09\u4e0d\u7a33\u5b9a\uff0c\u5efa\u8bae\u4e0e\u5916\u5305\u65b9\u7ea6\u5b9a\u660e\u786e\u7684\u54cd\u5e94\u65f6\u9650\u5e76\u7eb3\u5165\u5408\u540c\u8003\u6838\u3002", { }),
  body("\u4e8c\u662f\u6237\u5916\u8bbe\u65bd\u7ef4\u4fee\u6210\u672c\u504f\u9ad8\u3002\u673a\u68b0\u635f\u574f\u7c7b\u8d39\u7528\u5360\u6bd4\u63a5\u8fd1\u56db\u6210\uff0c\u5efa\u8bae\u5728\u6237\u5916\u8bbe\u65bd\u5468\u8fb9\u589e\u8bbe\u9632\u62a4\u63d0\u793a\uff0c\u5e76\u5c06\u6bcf\u6708\u4e00\u6b21\u7684\u6237\u5916\u8bbe\u65bd\u5b89\u5168\u5de1\u67e5\u7eb3\u5165\u4fdd\u5b89\u4f8b\u884c\u5de5\u4f5c\u3002"),
  body("\u4e09\u662f\u90e8\u5206\u8bbe\u5907\u8d85\u671f\u670d\u5f79\u3002\u5f85\u62a5\u5e9f 2 \u53f0\u3001\u95f2\u7f6e 3 \u53f0\u957f\u671f\u5360\u7528\u5e93\u623f\u4e0e\u6559\u5ba4\u7a7a\u95f4\uff0c\u5e94\u5728 11 \u6708\u5185\u5b8c\u6210\u8d44\u4ea7\u5904\u7f6e\u5ba1\u6279\u4e0e\u8c03\u62e8\uff0c\u907f\u514d\u8d44\u4ea7\u95f2\u7f6e\u635f\u5931\u3002"),

  h1("\u4e94\u300111 \u6708\u5de5\u4f5c\u8ba1\u5212"),
  body("1. \u529e\u7ed3 10 \u6708\u7ed3\u8f6c\u7ef4\u4fee\u5355 3 \u5355\uff0c\u91cd\u70b9\u8ddf\u8fdb SB-004 \u53f0\u5f0f\u7535\u8111\u6d3e\u5355\uff0c\u76ee\u6807 11 \u6708 10 \u65e5\u524d\u5168\u90e8\u5b8c\u6210\uff1b"),
  body("2. \u5b8c\u6210\u5f85\u62a5\u5e9f\u8bbe\u5907\uff08SB-014\u3001SB-027\uff09\u7684\u8d44\u4ea7\u5904\u7f6e\u7533\u62a5\uff0c\u542f\u52a8 3 \u53f0\u95f2\u7f6e\u8bbe\u5907\u7684\u8de8\u56ed\u8c03\u62e8\u8bc4\u4f30\uff1b"),
  body("3. \u5165\u51ac\u524d\u5b8c\u6210\u5168\u90e8\u7a7a\u8c03\u8bbe\u5907\u96c6\u4e2d\u4fdd\u517b\uff0c\u5f62\u6210\u4fdd\u517b\u8bb0\u5f55\u5e76\u5f55\u5165\u7cfb\u7edf\uff1b"),
  body("4. \u7ec4\u7ec7\u5404\u56ed\u4fdd\u7ba1\u4eba\u5f00\u5c55\u300a\u4fe1\u606f\u8bbe\u5907\u4f7f\u7528\u4e0e\u7ef4\u62a4\u300b\u57f9\u8bad\uff0c\u5ba3\u8d2f\u300aExcel \u6570\u636e\u5f55\u5165\u89c4\u8303\u300b\uff0c\u7edf\u4e00\u5f55\u5165\u53e3\u5f84\uff1b"),
  body("5. \u6309\u6708\u5237\u65b0\u7ba1\u7406\u770b\u677f\u6570\u636e\uff0c\u4e8e\u6708\u521d\u4e09\u4e2a\u5de5\u4f5c\u65e5\u5185\u53d1\u5e03\u4e0a\u6708\u6708\u5ea6\u62a5\u544a\u3002"),

  h1("\u9644\uff1a\u6570\u636e\u6765\u6e90\u4e0e\u7f16\u5236\u8bf4\u660e"),
  body("\u672c\u62a5\u544a\u5168\u90e8\u7edf\u8ba1\u6570\u636e\u5747\u53d6\u81ea\u300a\u56ed\u52a1\u8bbe\u5907\u7ba1\u7406\u7cfb\u7edf.xlsx\u300b\u3002\u8be5\u5de5\u4f5c\u7c3f\u5305\u542b\u8bbe\u5907\u53f0\u8d26\u3001\u7ef4\u4fee\u8bb0\u5f55\u3001\u6708\u5ea6\u7edf\u8ba1\u3001\u7ba1\u7406\u770b\u677f\u3001\u6570\u636e\u900f\u89c6\u4e94\u4e2a\u5de5\u4f5c\u8868\uff0c\u7edf\u8ba1\u6307\u6807\u7531 COUNTIF\u3001SUMIFS \u7b49\u516c\u5f0f\u5bf9\u6e90\u6570\u636e\u5b9e\u65f6\u6c47\u603b\uff0c\u6570\u636e\u5f55\u5165\u5747\u91c7\u7528\u4e0b\u62c9\u5217\u8868\u6807\u51c6\u5316\u63a7\u5236\u3002\u5982\u5bf9\u6570\u636e\u53e3\u5f84\u6709\u7591\u95ee\uff0c\u8bf7\u8054\u7cfb\u96c6\u56e2\u4fe1\u606f\u7ba1\u7406\u5458\u3002"),
];

const doc = new Document({
  creator: "Z.ai",
  styles: {
    default: {
      document: {
        run: { font: { ascii: "Times New Roman", eastAsia: "SimSun" }, size: 24, color: "000000" },
        paragraph: { spacing: { line: 312 } },
      },
    },
  },
  sections: [
    {
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 0, bottom: 0, left: 0, right: 0 } } },
      children: buildCoverR1({
        title: "\u56ed\u52a1\u8bbe\u5907\u4fe1\u606f\u5316\u7ba1\u7406\u6708\u5ea6\u5de5\u4f5c\u62a5\u544a",
        subtitle: "\u5e7c\u513f\u56ed\u96c6\u56e2\u529e\u516c\u81ea\u52a8\u5316\u4e0e\u6570\u636e\u62a5\u8868\u7cfb\u7edf \u00b7 2026 \u5e74 10 \u6708",
        englishLabel: "MONTHLY REPORT",
        metaLines: [
          "\u7f16\u5236\u90e8\u95e8\uff1a\u96c6\u56e2\u529e\u516c\u5ba4 \u4fe1\u606f\u7ba1\u7406\u5c97",
          "\u62a5\u544a\u671f\u95f4\uff1a2026 \u5e74 10 \u6708 1 \u65e5 \u81f3 10 \u6708 31 \u65e5",
          "\u6570\u636e\u6765\u6e90\uff1a\u56ed\u52a1\u8bbe\u5907\u7ba1\u7406\u7cfb\u7edf.xlsx",
          "\u7f16\u5236\u65e5\u671f\uff1a2026 \u5e74 10 \u6708 31 \u65e5",
        ],
        footerLeft: "\u5e7c\u513f\u56ed\u96c6\u56e2",
        footerRight: "\u5185\u90e8\u5de5\u4f5c\u6587\u4ef6",
        palette: P,
      }),
    },
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "606060" })],
          })],
        }),
      },
      children: bodyChildren,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(__dirname + "/\u6708\u5ea6\u4fe1\u606f\u5de5\u4f5c\u62a5\u544a_2026\u5e7410\u6708.docx", buf);
  console.log("saved docx");
});
