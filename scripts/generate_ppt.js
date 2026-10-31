// 幼儿园集团 — 信息设备使用与维护培训（12页）
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "Z.ai";
pres.title = "信息设备使用与维护培训";

const W = 13.33, H = 7.5, M = 0.6;
// 配色：暖青主色（与 Excel 看板同源）+ 橙色点缀
const BG = "FFFFFF";
const DARK = "0E3F3A";      // 封面/封底深青
const PRIMARY = "15857A";
const PRIMARY_L = "E4F2F0"; // 主色浅底
const ACCENT = "FF6A3B";
const TEXT = "1F2937";
const MUTED = "6B7280";
const FONT = "微软雅黑";

const bu = () => ({ code: "25B8", indent: 12 });

// ── 通用构件 ──
function pageHeader(slide, num, title, sub) {
  slide.background = { color: BG };
  slide.addText(String(num).padStart(2, "0"), {
    x: M, y: 0.42, w: 1.0, h: 0.7, margin: 0,
    fontSize: 30, bold: true, color: PRIMARY, fontFace: FONT,
  });
  slide.addText(title, {
    x: M + 1.05, y: 0.42, w: W - 2 * M - 1.05, h: 0.7, margin: 0,
    fontSize: 30, bold: true, color: TEXT, fontFace: FONT, valign: "middle",
  });
  if (sub) slide.addText(sub, {
    x: M + 1.05, y: 1.12, w: W - 2 * M - 1.05, h: 0.4, margin: 0,
    fontSize: 13, color: MUTED, fontFace: FONT,
  });
}
function sourceNote(slide, text) {
  slide.addText(text, {
    x: M, y: H - 0.55, w: W - 2 * M, h: 0.35, margin: 0,
    fontSize: 11, color: MUTED, fontFace: FONT, align: "left",
  });
}
// 行列表（图标方块 + 标题 + 描述）
function rowList(slide, rows, opts = {}) {
  const x = opts.x ?? M, w = opts.w ?? (W - 2 * M);
  const y0 = opts.y ?? 1.7, rh = opts.rh ?? 0.95;
  rows.forEach((r, i) => {
    const y = y0 + i * rh;
    slide.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.08, w: 0.34, h: 0.34, fill: { color: i % 2 ? ACCENT : PRIMARY } });
    slide.addText(r.n, { x, y: y + 0.08, w: 0.34, h: 0.34, margin: 0, fontSize: 14, bold: true, color: "FFFFFF", fontFace: FONT, align: "center", valign: "middle" });
    slide.addText(r.t, { x: x + 0.55, y, w: 3.6, h: 0.5, margin: 0, fontSize: 16, bold: true, color: TEXT, fontFace: FONT, valign: "middle" });
    slide.addText(r.d, { x: x + 4.2, y, w: w - 4.2, h: 0.5, margin: 0, fontSize: 13, color: MUTED, fontFace: FONT, valign: "middle" });
    if (i < rows.length - 1) slide.addShape(pres.shapes.LINE, { x: x + 0.55, y: y + rh - 0.14, w: w - 0.55, h: 0, line: { color: "E5E7EB", width: 0.75 } });
  });
}

// ═══ 1. 封面 ═══
let s = pres.addSlide();
s.background = { color: DARK };
s.addShape(pres.shapes.OVAL, { x: 9.2, y: -2.2, w: 6.5, h: 6.5, fill: { color: PRIMARY, transparency: 55 } });
s.addShape(pres.shapes.OVAL, { x: 11.4, y: 3.4, w: 3.6, h: 3.6, fill: { color: ACCENT, transparency: 35 } });
s.addText("幼儿园集团 · 员工内训", { x: M, y: 1.5, w: 8, h: 0.5, margin: 0, fontSize: 16, color: "9FD4CD", fontFace: FONT, charSpacing: 4 });
s.addText("信息设备使用与维护培训", { x: M, y: 2.1, w: 11, h: 1.3, margin: 0, fontSize: 52, bold: true, color: "FFFFFF", fontFace: FONT });
s.addText("规范使用 · 及时报修 · 数据留痕 —— 让每一台设备都有账可查", { x: M, y: 3.5, w: 10.5, h: 0.6, margin: 0, fontSize: 18, color: "C9E5E1", fontFace: FONT });
s.addText([
  { text: "主讲：集团信息管理员", options: { breakLine: true } },
  { text: "时间：2026 年 10 月", options: {} },
], { x: M, y: 5.6, w: 6, h: 1.0, margin: 0, fontSize: 14, color: "8FB8B2", fontFace: FONT, paraSpaceAfter: 6 });
s.addNotes("开场：说明培训对象为三园全体教职工与保育员，时长约40分钟，培训后需完成扫码小测。");

// ═══ 2. 培训目标与议程 ═══
s = pres.addSlide();
pageHeader(s, 1, "培训目标与议程", "三个目标、四个环节，先讲怎么用，再讲坏了怎么办，最后讲数据怎么录");
const goals = [
  { n: "01", t: "会使用", d: "掌握一体机、消毒柜、空调等常用设备的规范操作与日常保养" },
  { n: "02", t: "会报修", d: "记住一条报修路径：设备编号 + 故障现象，5 分钟完成报修登记" },
  { n: "03", t: "会录入", d: "按《Excel 数据录入规范》录入台账与维修记录，数据口径全园统一" },
];
goals.forEach((g, i) => {
  const x = M + i * 4.11;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.8, w: 3.8, h: 2.1, fill: { color: PRIMARY_L }, rectRadius: 0.08 });
  s.addText(g.n, { x: x + 0.3, y: 2.0, w: 1.4, h: 0.8, margin: 0, fontSize: 40, bold: true, color: PRIMARY, fontFace: FONT });
  s.addText(g.t, { x: x + 0.3, y: 2.75, w: 3.2, h: 0.5, margin: 0, fontSize: 19, bold: true, color: TEXT, fontFace: FONT });
  s.addText(g.d, { x: x + 0.3, y: 3.2, w: 3.2, h: 0.6, margin: 0, fontSize: 12.5, color: MUTED, fontFace: FONT });
});
const agenda = [
  { n: "1", t: "设备家底与系统介绍", d: "10 分钟" },
  { n: "2", t: "规范使用与日常保养", d: "10 分钟" },
  { n: "3", t: "报修流程与常见故障", d: "12 分钟" },
  { n: "4", t: "数据录入规范要点", d: "8 分钟" },
];
agenda.forEach((a, i) => {
  const x = M + i * 3.13;
  s.addShape(pres.shapes.OVAL, { x, y: 4.7, w: 0.62, h: 0.62, fill: { color: i === 3 ? ACCENT : PRIMARY } });
  s.addText(a.n, { x, y: 4.7, w: 0.62, h: 0.62, margin: 0, fontSize: 18, bold: true, color: "FFFFFF", fontFace: FONT, align: "center", valign: "middle" });
  s.addText(a.t, { x, y: 5.45, w: 2.9, h: 0.75, margin: 0, fontSize: 14, bold: true, color: TEXT, fontFace: FONT });
  s.addText(a.d, { x, y: 6.15, w: 2.9, h: 0.4, margin: 0, fontSize: 12, color: MUTED, fontFace: FONT });
});

// ═══ 3. 设备家底 ═══
s = pres.addSlide();
pageHeader(s, 2, "先认识我们的设备家底", "数据来自《园务设备管理系统.xlsx》，每月更新一次");
const stats = [
  { v: "32", u: "台", l: "在册设备总量" },
  { v: "3", u: "所园", l: "阳光 / 彩虹 / 星辰" },
  { v: "20", u: "单", l: "近5个月累计报修" },
  { v: "6,740", u: "元", l: "累计维修费用" },
];
stats.forEach((t, i) => {
  const x = M + i * 3.13;
  s.addText([
    { text: t.v, options: { fontSize: 44, bold: true, color: i === 3 ? ACCENT : PRIMARY } },
    { text: " " + t.u, options: { fontSize: 15, color: MUTED } },
  ], { x, y: 1.55, w: 2.9, h: 0.85, margin: 0, fontFace: FONT });
  s.addText(t.l, { x, y: 2.4, w: 2.9, h: 0.4, margin: 0, fontSize: 12.5, color: MUTED, fontFace: FONT });
});
s.addChart(pres.charts.BAR, [
  { name: "设备总数", labels: ["阳光园", "彩虹园", "星辰园"], values: [12, 11, 9] },
  { name: "其中维修中", labels: ["阳光园", "彩虹园", "星辰园"], values: [2, 1, 1] },
], {
  x: M, y: 3.0, w: 7.4, h: 3.7, barDir: "col",
  chartColors: [PRIMARY, ACCENT],
  chartArea: { fill: { color: "FFFFFF" } },
  catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
  catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
  valGridLine: { color: "E5E7EB", size: 0.5 }, catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TEXT, dataLabelFontSize: 11, dataLabelFontFace: FONT,
  showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: FONT, legendColor: MUTED,
  barGapWidthPct: 60,
});
s.addText([
  { text: "状态速览：", options: { bold: true, color: TEXT, breakLine: true } },
  { text: "在用 23 台（71.9%）", options: { color: TEXT, breakLine: true } },
  { text: "维修中 4 台 · 闲置 3 台", options: { color: TEXT, breakLine: true } },
  { text: "待报废 2 台（投影仪、针打）", options: { color: TEXT, breakLine: true } },
  { text: "在用率 71.9%，整体状况良好；每台设备都有唯一编号，认准编号再报修。", options: { color: MUTED } },
], { x: 8.4, y: 3.2, w: 4.3, h: 3.2, margin: 0, fontSize: 14, fontFace: FONT, paraSpaceAfter: 8, valign: "top" });
sourceNote(s, "数据来源：园务设备管理系统.xlsx（月度统计表，截至 2026-10-31）");

// ═══ 4. 系统五张表 ═══
s = pres.addSlide();
pageHeader(s, 3, "一个工作簿，五张表", "所有信息都录在同一个 Excel 里，各表自动联动，不需要手工再统计");
rowList(s, [
  { n: "1", t: "设备台账", d: "32 台设备的户口本：编号、名称、园区、采购日期、原值、使用状态" },
  { n: "2", t: "维修记录", d: "每一单报修的病历：设备编号带出名称园区，报修日期、故障类型、费用" },
  { n: "3", t: "月度统计", d: "COUNTIF / SUMIFS 公式自动汇总，园区概况、故障分布、月度趋势" },
  { n: "4", t: "管理看板", d: "5 个 KPI + 4 张图表，源数据更新后看板自动刷新" },
  { n: "5", t: "数据透视", d: "按园区 × 故障类型交叉汇总报修单量与费用，供月度分析" },
], { y: 1.75, rh: 1.02 });

// ═══ 5. 规范使用 ═══
s = pres.addSlide();
pageHeader(s, 4, "设备规范使用四条要求", "大部分故障源于使用不当——做到这四条，可减少约一半报修");
rowList(s, [
  { n: "1", t: "对号使用", d: "使用前核对设备编号与登记位置一致，发现设备换位先问信息员" },
  { n: "2", t: "规范开关", d: "一体机、投影仪按顺序开关机；消毒柜高温程序结束后等待 10 分钟再取放" },
  { n: "3", t: "清洁消毒", d: "幼儿可触碰表面每日清水擦拭，每周按园所消毒计划执行并记录" },
  { n: "4", t: "移动报备", d: "设备不跨园、不跨班挪用；确需调整的，报信息员在台账中同步变更" },
], { y: 1.75, rh: 1.02 });
s.addNotes("强调：设备移动不报备是台账错乱的主要原因，去年因此产生过3次找不到设备的记录。");

// ═══ 6. 报修流程 ═══
s = pres.addSlide();
pageHeader(s, 5, "报修流程：五步走", "从发现故障到验收销单，全程在维修记录表里留痕");
const steps = ["发现故障", "填写报修单", "信息员派单", "维修实施", "验收销单"];
const stepDesc = [
  "先断电、隔离，避免故障扩大或伤到幼儿",
  "维修记录表新增一行：编号+故障类型+现象",
  "信息员确认类型，联系维修方，状态改为维修中",
  "维修完成后补录完成日期与维修费用",
  "报修人试用确认，状态改为已完成",
];
steps.forEach((t, i) => {
  const x = M + i * 2.55;
  const y = 2.3;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.25, h: 1.5, fill: { color: i === 1 ? ACCENT : PRIMARY }, rectRadius: 0.07 });
  s.addText(String(i + 1), { x: x + 0.15, y: y + 0.12, w: 0.9, h: 0.45, margin: 0, fontSize: 20, bold: true, color: "FFFFFF", fontFace: FONT });
  s.addText(t, { x: x + 0.15, y: y + 0.62, w: 2.0, h: 0.7, margin: 0, fontSize: 15, bold: true, color: "FFFFFF", fontFace: FONT });
  s.addText(stepDesc[i], { x: x + 0.05, y: y + 1.7, w: 2.15, h: 1.3, margin: 0, fontSize: 11.5, color: MUTED, fontFace: FONT });
  if (i < steps.length - 1) {
    s.addShape(pres.shapes.RECTANGLE, { x: x + 2.28, y: y + 0.62, w: 0.24, h: 0.26, fill: { color: "D1D5DB" } });
  }
});
s.addText([
  { text: "关键提醒：", options: { bold: true, color: ACCENT } },
  { text: "报修只认设备编号。编号在设备左上角的银色标签上；找不到编号的设备直接联系信息员处理。", options: { color: TEXT } },
], { x: M, y: 5.9, w: W - 2 * M, h: 0.7, margin: 0, fontSize: 14, fontFace: FONT });

// ═══ 7. 常见故障 ═══
s = pres.addSlide();
pageHeader(s, 6, "常见故障：哪些能自己处理，哪些立即报修", "左列自查复位，右列一律停用并报修——判断不准时一律报修");
const th = (t) => ({ text: t, options: { fill: { color: PRIMARY }, color: "FFFFFF", bold: true, fontFace: FONT, fontSize: 13.5, align: "center", valign: "middle" } });
const td = (t, opts = {}) => ({ text: t, options: { fontFace: FONT, fontSize: 12.5, color: TEXT, valign: "middle", ...opts } });
s.addTable([
  [th("故障现象"), th("先尝试"), th("仍无效"), th("报修类型")],
  [td("一体机触控失灵 / 卡顿"), td("重启设备，检查连线"), td("立即报修", { bold: true, color: ACCENT }), td("触控失灵 / 软件故障")],
  [td("空调不制冷 / 制热弱"), td("检查遥控模式、滤网清洁"), td("立即报修", { bold: true, color: ACCENT }), td("制冷异常")],
  [td("打印机卡纸"), td("按走纸方向取纸，换纸重试"), td("立即报修", { bold: true, color: ACCENT }), td("进纸卡纸")],
  [td("消毒柜 / 冰箱不通电"), td("检查插座与漏保开关"), td("立即报修", { bold: true, color: ACCENT }), td("电路故障")],
  [td("户外滑梯、攀爬架松动"), td("不停用自查，设围挡"), td("立即报修", { bold: true, color: ACCENT }), td("机械损坏")],
], {
  x: M, y: 1.75, w: W - 2 * M, colW: [3.4, 3.3, 1.9, 2.91],
  border: { pt: 0.75, color: "E5E7EB" }, rowH: 0.62,
  fill: { color: "FFFFFF" }, align: "left", margin: 0.08,
});
s.addText("近5个月机械损坏类维修费 2,480 元，占全部维修费 36.8%——户外设施发现问题请第一时间围挡报修，早报早修更省钱。",
  { x: M, y: 6.15, w: W - 2 * M, h: 0.6, margin: 0, fontSize: 13, color: MUTED, fontFace: FONT });
sourceNote(s, "数据来源：园务设备管理系统.xlsx · 故障类型分布（2026-06 至 2026-10）");

// ═══ 8. 故障分布与趋势 ═══
s = pres.addSlide();
pageHeader(s, 7, "近5个月故障分布与维修趋势", "两张图来自管理看板，每月初随月度报告同步更新");
s.addChart(pres.charts.DOUGHNUT, [
  { name: "报修单量", labels: ["机械损坏", "制冷异常", "电路故障", "软件故障", "触控失灵", "耗材老化", "其他"], values: [3, 3, 3, 3, 2, 2, 4] },
], {
  x: M, y: 1.8, w: 5.9, h: 4.6, holeSize: 55,
  chartColors: [PRIMARY, "1E9E90", "2FA79A", "10705F", ACCENT, "D95423", "7D828A"],
  showPercent: true, dataLabelColor: "FFFFFF", dataLabelFontSize: 11, dataLabelFontFace: FONT,
  showLegend: true, legendPos: "r", legendFontSize: 12, legendFontFace: FONT, legendColor: TEXT,
  showTitle: true, title: "故障类型分布（报修单量 20 单）", titleFontSize: 14, titleColor: TEXT, titleFontFace: FONT,
});
s.addChart(pres.charts.LINE, [
  { name: "报修数量", labels: ["6月", "7月", "8月", "9月", "10月"], values: [3, 4, 4, 5, 4] },
  { name: "已完成", labels: ["6月", "7月", "8月", "9月", "10月"], values: [3, 4, 4, 3, 1] },
], {
  x: 6.9, y: 1.8, w: 5.8, h: 4.6,
  chartColors: [PRIMARY, ACCENT], lineSize: 2.5, lineSmooth: false,
  lineDataSymbol: "circle", lineDataSymbolSize: 7,
  catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
  catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
  valGridLine: { color: "E5E7EB", size: 0.5 }, catGridLine: { style: "none" },
  showValue: true, dataLabelColor: TEXT, dataLabelFontSize: 11, dataLabelFontFace: FONT,
  showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: FONT, legendColor: MUTED,
  showTitle: true, title: "月度报修与完成趋势（累计完成率 75%）", titleFontSize: 14, titleColor: TEXT, titleFontFace: FONT,
});
sourceNote(s, "数据来源：园务设备管理系统.xlsx · 管理看板（2026-06 至 2026-10，报修 20 单 / 完成 15 单 / 费用 6,740 元）");

// ═══ 9. 数据录入规范 ═══
s = pres.addSlide();
pageHeader(s, 8, "数据录入规范：四条铁律", "详细规则见《Excel 数据录入规范》，这里记最关键的四条");
rowList(s, [
  { n: "1", t: "只用下拉列表", d: "使用状态、故障类型、维修状态必须从下拉选择，禁止手工输入变体说法" },
  { n: "2", t: "编号是唯一钥匙", d: "维修记录只填设备编号，名称与园区由 VLOOKUP 自动带出，禁止手写" },
  { n: "3", t: "当日事当日记", d: "报修、维修完成、设备变更当天录入，月度统计以录入日期为准" },
  { n: "4", t: "不碰公式区", d: "月度统计、管理看板、数据透视由公式生成，只读不改，发现异常找信息员" },
], { y: 1.75, rh: 1.02 });

// ═══ 10. 日常维护日历 ═══
s = pres.addSlide();
pageHeader(s, 9, "日常维护日历", "三档频次，责任到人，做完在台账备注栏打钩留痕");
const cal = [
  { n: "日", t: "每日", d: "教学设备关机断电；厨房电器清洁归位；户外设施巡查一圈", who: "当班教师 / 厨房 / 保安" },
  { n: "周", t: "每周", d: "设备表面消毒并记录；空调滤网目视检查；打印耗材余量检查", who: "各班保育员" },
  { n: "季", t: "每季", d: "空调深度保养；户外设施紧固件检查；台账与实物盘点核对", who: "信息员 + 后勤主管" },
];
cal.forEach((c, i) => {
  const y = 1.85 + i * 1.55;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: 1.5, h: 1.3, fill: { color: i === 2 ? ACCENT : PRIMARY }, rectRadius: 0.08 });
  s.addText(c.n, { x: M, y, w: 1.5, h: 1.3, margin: 0, fontSize: 34, bold: true, color: "FFFFFF", fontFace: FONT, align: "center", valign: "middle" });
  s.addText(c.t + "　" + c.d, { x: M + 1.8, y: y + 0.08, w: 8.6, h: 0.7, margin: 0, fontSize: 15, bold: true, color: TEXT, fontFace: FONT, valign: "middle" });
  s.addText("责任人：" + c.who, { x: M + 1.8, y: y + 0.78, w: 8.6, h: 0.45, margin: 0, fontSize: 12.5, color: MUTED, fontFace: FONT });
  if (i < cal.length - 1) s.addShape(pres.shapes.LINE, { x: M + 1.8, y: y + 1.42, w: W - 2 * M - 1.8, h: 0, line: { color: "E5E7EB", width: 0.75 } });
});

// ═══ 11. 安全注意事项 ═══
s = pres.addSlide();
pageHeader(s, 10, "安全三条底线", "设备安全的前提是不出安全事故");
const safes = [
  { n: "用电", d: "插座加装防护盖；电器检修必须先断电；移动电器先拔插头再挪动", big: "断电再动" },
  { n: "儿童", d: "幼儿活动区域内，一体机支架、消毒柜门、滑梯连接件每周检查一次", big: "隔离检查" },
  { n: "数据", d: "台账与维修记录不得外传；含幼儿信息的截图与表格不入个人手机相册", big: "不外传" },
];
safes.forEach((c, i) => {
  const x = M + i * 4.11;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.9, w: 3.8, h: 3.6, fill: { color: PRIMARY_L }, rectRadius: 0.08 });
  s.addText(c.n, { x: x + 0.35, y: 2.2, w: 3.1, h: 0.55, margin: 0, fontSize: 21, bold: true, color: PRIMARY, fontFace: FONT });
  s.addText(c.big, { x: x + 0.35, y: 2.85, w: 3.1, h: 0.75, margin: 0, fontSize: 27, bold: true, color: ACCENT, fontFace: FONT });
  s.addText(c.d, { x: x + 0.35, y: 3.75, w: 3.1, h: 1.5, margin: 0, fontSize: 13, color: TEXT, fontFace: FONT, lineSpacingMultiple: 1.25 });
});

// ═══ 12. 结束页 ═══
s = pres.addSlide();
s.background = { color: DARK };
s.addShape(pres.shapes.OVAL, { x: -2.5, y: 3.8, w: 6.5, h: 6.5, fill: { color: PRIMARY, transparency: 55 } });
s.addText("谢谢参与", { x: M, y: 2.5, w: 12.13, h: 1.2, margin: 0, fontSize: 54, bold: true, color: "FFFFFF", fontFace: FONT, align: "center" });
s.addText("设备有问题，找信息员：园区办公室内线 8008", { x: M, y: 3.9, w: 12.13, h: 0.6, margin: 0, fontSize: 18, color: "C9E5E1", fontFace: FONT, align: "center" });
s.addText("配套文件：《园务设备管理系统.xlsx》《月度信息工作报告》《Excel 数据录入规范》",
  { x: M, y: 4.7, w: 12.13, h: 0.5, margin: 0, fontSize: 13, color: "8FB8B2", fontFace: FONT, align: "center" });

pres.writeFile({ fileName: __dirname + "/信息设备使用与维护培训.pptx" }).then(() => console.log("saved pptx"));
