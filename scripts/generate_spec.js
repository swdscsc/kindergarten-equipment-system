// 《Excel 数据录入规范》— 园务设备管理系统配套制度文档
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  TableLayoutType, Footer, PageNumber, NumberFormat,
} = require("docx");
const fs = require("fs");

const TEAL = "15857A";
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160, line: 312 },
    children: [new TextRun({ text, bold: true, size: 32, color: "000000", font: { ascii: "Times New Roman", eastAsia: "SimHei" } })],
  });
}
function body(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 480 },
    spacing: { line: 312, after: 80 },
    children: [new TextRun({ text, size: 24, color: "000000", font: { ascii: "Times New Roman", eastAsia: "SimSun" } })],
  });
}
function tableTitle(text) {
  return new Paragraph({
    keepNext: true, spacing: { before: 200, after: 100, line: 312 },
    children: [new TextRun({ text, bold: true, size: 21, color: "000000", font: { ascii: "Times New Roman", eastAsia: "SimHei" } })],
  });
}
function cellPara(text, header = false, center = false) {
  return new Paragraph({
    alignment: header || center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { line: 276 },
    children: [new TextRun({ text: String(text), size: 21, bold: header,
      color: header ? "FFFFFF" : "000000",
      font: { ascii: "Times New Roman", eastAsia: header ? "SimHei" : "SimSun" } })],
  });
}
function dataTable(headers, rows, widths, centerCols = []) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: TEAL },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL },
      left: NB, right: NB,
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D5D0C8" },
      insideVertical: NB,
    },
    rows: [
      new TableRow({
        tableHeader: true, cantSplit: true,
        children: headers.map((h, i) => new TableCell({
          width: { size: widths[i], type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: TEAL },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [cellPara(h, true)],
        })),
      }),
      ...rows.map((r, ri) => new TableRow({
        cantSplit: true,
        children: r.map((v, i) => new TableCell({
          width: { size: widths[i], type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: ri % 2 === 1 ? "F0EDE5" : "FFFFFF" },
          margins: { top: 70, bottom: 70, left: 120, right: 120 },
          children: [cellPara(v, false, centerCols.includes(i))],
        })),
      })),
    ],
  });
}

const children = [
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 200, after: 120, line: Math.ceil(22 * 23), lineRule: "atLeast" },
    children: [new TextRun({ text: "园务设备管理系统 Excel 数据录入规范", bold: true, size: 44, color: "000000", font: { ascii: "Times New Roman", eastAsia: "SimHei" } })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 300, line: 312 },
    children: [new TextRun({ text: "（2026 年 10 月修订 · 幼儿园集团）", size: 24, color: "404040", font: { ascii: "Times New Roman", eastAsia: "SimSun" } })],
  }),
  tableTitle("表 0  文件信息"),
  dataTable(["文件编号", "适用范围", "生效日期", "修订周期"],
    [["KG-XX-2026-01", "三园全体教职工及信息管理员", "2026-11-01", "每学期复核一次"]],
    [25, 35, 20, 20]),

  h1("一、目的与适用范围"),
  body("为统一《园务设备管理系统.xlsx》的数据录入口径，保证设备台账、维修记录、月度统计、管理看板与数据透视各工作表数据准确、口径一致、可追溯，特制定本规范。本规范适用于阳光园、彩虹园、星辰园三所幼儿园的全体教职工，信息管理员负责执行监督。"),

  h1("二、岗位职责与权限"),
  tableTitle("表 1  岗位分工"),
  dataTable(["岗位", "职责", "系统权限"],
    [
      ["各班教师 / 保育员", "报告设备故障，填写报修记录行", "维修记录表：新增行"],
      ["厨房 / 后勤人员", "报告厨房电器故障，执行每日清洁", "维修记录表：新增行"],
      ["园区信息管理员", "派单、核验、状态维护、月末统计发布", "全部工作表（统计区只读）"],
      ["集团后勤主管", "审批设备处置与跨园调拨", "设备台账：状态变更审批"],
    ],
    [22, 48, 30]),

  h1("三、工作簿结构与字段规范"),
  body("工作簿共五个工作表，录入人员只允许在「设备台账」「维修记录」两个工作表的录入区新增或修改数据；「月度统计」「管理看板」「数据透视」为公式自动生成区，任何人员不得手工改写。"),
  tableTitle("表 2  设备台账字段规范"),
  dataTable(["字段", "格式要求", "录入方式", "必填"],
    [
      ["设备编号", "SB-XXX，三园唯一，终身不变", "新增设备时由信息管理员统一编号", "是"],
      ["设备名称 / 类别", "按采购合同名称填写", "下拉列表或规范名称", "是"],
      ["园区", "阳光园 / 彩虹园 / 星辰园", "下拉列表选择", "是"],
      ["采购日期", "yyyy-mm-dd 日期格式", "日期选择，禁止文本录入", "是"],
      ["采购原值", "货币格式，保留两位小数", "数字录入，不含「元」字", "是"],
      ["已用年限", "公式自动计算（DATEDIF）", "禁止手工填写", "自动"],
      ["使用状态", "在用 / 维修中 / 闲置 / 待报废", "下拉列表选择", "是"],
      ["保管人", "教职工真实姓名", "文本录入", "是"],
    ],
    [20, 34, 34, 12], [3]),
  tableTitle("表 3  维修记录字段规范"),
  dataTable(["字段", "格式要求", "录入方式", "必填"],
    [
      ["报修单号", "WX-YYMM-XX（例：WX-2610-01），当月顺序编号", "按规则手工编号，不得跳号重号", "是"],
      ["设备编号", "必须与台账一致", "文本录入；名称与园区由 VLOOKUP 自动带出", "是"],
      ["故障类型", "标准词表：机械损坏、制冷异常、电路故障、软件故障、耗材老化、触控失灵、管路漏水、光学部件、进纸卡纸、噪音异常", "下拉列表选择；新增类型由信息管理员维护词表", "是"],
      ["报修日期 / 完成日期", "yyyy-mm-dd 日期格式", "日期选择；未完成时完成日期留空", "是"],
      ["维修状态", "待派单 / 维修中 / 已完成", "下拉列表选择", "是"],
      ["维修费用", "货币格式；待派单阶段填 0", "数字录入", "是"],
      ["处理天数", "公式自动计算（完成日期-报修日期）", "禁止手工填写", "自动"],
    ],
    [20, 34, 34, 12], [3]),

  h1("四、录入操作规则"),
  body("（一）只用下拉列表。所有含下拉列表的字段一律从列表中选择，禁止手工输入「坏了」「修好了」等描述性文字替代标准状态值；下拉词表由信息管理员统一维护，个人不得自行增删。"),
  body("（二）编号是唯一钥匙。维修记录中只填写设备编号，设备名称与园区由公式自动带出；发现带出结果为「—」时，说明编号不存在或已被改动，应立即核对，不得手工覆盖公式。"),
  body("（三）当日事当日记。报修、派单、维修完成、状态变更应当在发生当日录入；月度统计与看板以录入数据实时汇总，补录需在备注中说明原因。"),
  body("（四）不碰公式区。月度统计、管理看板、数据透视三个工作表全部由公式与透视缓存生成，为只读区域；发现统计异常应联系信息管理员从源数据排查，禁止直接改写统计结果。"),
  body("（五）日期与金额格式统一。日期一律使用 yyyy-mm-dd 格式的真日期（可参与计算），金额以元为单位保留两位小数，禁止使用「10月5日」「八百元」等文本写法。"),

  h1("五、数据修改与更正"),
  body("已录入数据的更正遵循「谁录入、谁更正、留痕迹」原则：直接在原单元格修改并保留正确值即可，系统按最终值统计；对已完成维修单的费用或日期调整，须经园区信息管理员确认后修改。跨月更正需同步检查当月统计口径是否受影响，必要时在月度报告中说明。"),

  h1("六、数据安全与备份"),
  body("工作簿存储于集团办公网盘指定目录，禁止通过个人微信、个人邮箱外发；含设备原值与维修费用的数据仅限本系统使用。信息管理员每周五下班前备份一次，文件名追加日期后缀（例：园务设备管理系统_20261031.xlsx），保留最近 8 份备份。学期末将全年数据导出 PDF 归档。"),

  h1("七、附则"),
  body("本规范自 2026 年 11 月 1 日起施行，由集团办公室负责解释。执行中的问题与改进建议请反馈至集团信息管理员；每学期末结合培训（见《信息设备使用与维护培训》课件）对本规范执行情况进行检查通报。"),
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
  sections: [{
    properties: {
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
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(__dirname + "/Excel\u6570\u636e\u5f55\u5165\u89c4\u8303.docx", buf);
  console.log("saved spec");
});
