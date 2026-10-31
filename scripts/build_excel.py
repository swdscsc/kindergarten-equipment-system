# -*- coding: utf-8 -*-
"""幼儿园集团办公自动化与数据报表系统 — Excel 工作簿构建脚本"""
import sys, os
from datetime import date

XLSX_SKILL_DIR = r"C:\Users\asus\.zcode\cli\plugins\cache\zcode-plugins-official\spreadsheets\0.2.1\skills\xlsx"
for sub in [XLSX_SKILL_DIR, os.path.join(XLSX_SKILL_DIR, "templates")]:
    if sub not in sys.path:
        sys.path.insert(0, sub)

import base
# 交付对象是用户 Windows 上的 Excel/WPS，使用系统自带的微软雅黑
base.FONT_NAME = "微软雅黑"
base.HEADER_BOLD = True

from base import (setup_sheet, style_header_row, style_data_row, style_total_row,
                  auto_fit_columns, auto_fit_row_heights, font_body, font_subheader,
                  font_caption, font_kpi, font_kpi_label, align_number, align_date,
                  make_chart_title, setup_chart_titles, apply_chart_colors,
                  create_bar_chart, create_line_chart, create_pie_chart)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import Reference

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "园务设备管理系统.xlsx")

FMT_INT = '#,##0'
FMT_PCT = '0.0%'
FMT_CNY = '¥#,##0.00'
FMT_DATE = 'yyyy-mm-dd'

PRIMARY = base.PRIMARY
ACC_POS, ACC_NEG, ACC_WARN = base.ACCENT_POSITIVE, base.ACCENT_NEGATIVE, base.ACCENT_WARNING
N900, N600, N100, N0 = base.NEUTRAL_900, base.NEUTRAL_600, base.NEUTRAL_100, base.NEUTRAL_0

wb = Workbook()

# ============================================================
# 工作表 1：设备台账
# ============================================================
ws1 = wb.active
ws1.title = "设备台账"
H1 = ["设备编号", "设备名称", "类别", "园区", "采购日期", "采购原值(元)", "已用年限", "使用状态", "保管人"]
LAST1 = 1 + len(H1)  # J
setup_sheet(ws1, title="幼儿园集团设备台账（截至 2026 年 10 月）", last_col=LAST1)
for c, h in enumerate(H1, 2):
    ws1.cell(row=4, column=c, value=h)
style_header_row(ws1, 4, 2, LAST1)

devices = [
    # 编号, 名称, 类别, 园区, 采购日期, 原值, 状态, 保管人
    ("SB-001", "教学一体机", "教学设备", "阳光园", date(2022,3,15), 12800, "在用", "张莉"),
    ("SB-002", "多媒体投影仪", "教学设备", "阳光园", date(2021,9,10), 4600, "在用", "张莉"),
    ("SB-003", "台式电脑", "办公设备", "阳光园", date(2023,6,20), 4200, "在用", "王芳"),
    ("SB-004", "台式电脑", "办公设备", "阳光园", date(2023,6,20), 4200, "维修中", "王芳"),
    ("SB-005", "激光打印机", "办公设备", "阳光园", date(2022,11,5), 2300, "在用", "王芳"),
    ("SB-006", "立柜式空调", "空调暖通", "阳光园", date(2020,5,18), 6800, "维修中", "李娜"),
    ("SB-007", "壁挂式空调", "空调暖通", "阳光园", date(2021,4,22), 3500, "在用", "李娜"),
    ("SB-008", "餐具消毒柜", "厨房电器", "阳光园", date(2022,8,30), 3200, "在用", "赵敏"),
    ("SB-009", "厨房冷藏冰箱", "厨房电器", "阳光园", date(2023,2,14), 4500, "在用", "赵敏"),
    ("SB-010", "户外组合滑梯", "户外设施", "阳光园", date(2019,7,8), 26800, "在用", "陈静"),
    ("SB-011", "户外攀爬架", "户外设施", "阳光园", date(2019,7,8), 9600, "闲置", "陈静"),
    ("SB-012", "幼儿晨检机", "教学设备", "阳光园", date(2024,3,12), 5600, "在用", "周燕"),
    ("SB-013", "教学一体机", "教学设备", "彩虹园", date(2022,3,15), 12800, "在用", "张莉"),
    ("SB-014", "多媒体投影仪", "教学设备", "彩虹园", date(2020,10,25), 4600, "待报废", "张莉"),
    ("SB-015", "台式电脑", "办公设备", "彩虹园", date(2023,6,20), 4200, "在用", "王芳"),
    ("SB-016", "激光打印机", "办公设备", "彩虹园", date(2021,12,1), 2300, "维修中", "王芳"),
    ("SB-017", "复印一体机", "办公设备", "彩虹园", date(2022,4,16), 8900, "在用", "王芳"),
    ("SB-018", "立柜式空调", "空调暖通", "彩虹园", date(2020,5,18), 6800, "在用", "李娜"),
    ("SB-019", "壁挂式空调", "空调暖通", "彩虹园", date(2022,6,9), 3500, "闲置", "李娜"),
    ("SB-020", "餐具消毒柜", "厨房电器", "彩虹园", date(2021,8,27), 3200, "在用", "赵敏"),
    ("SB-021", "和面机", "厨房电器", "彩虹园", date(2023,9,6), 2800, "在用", "赵敏"),
    ("SB-022", "户外组合滑梯", "户外设施", "彩虹园", date(2020,4,11), 26800, "在用", "陈静"),
    ("SB-023", "幼儿晨检机", "教学设备", "彩虹园", date(2024,3,12), 5600, "在用", "周燕"),
    ("SB-024", "触控一体机", "教学设备", "星辰园", date(2024,9,1), 15600, "在用", "张莉"),
    ("SB-025", "多媒体投影仪", "教学设备", "星辰园", date(2021,5,19), 4600, "在用", "张莉"),
    ("SB-026", "台式电脑", "办公设备", "星辰园", date(2023,6,20), 4200, "在用", "王芳"),
    ("SB-027", "针式打印机", "办公设备", "星辰园", date(2019,11,13), 1800, "待报废", "王芳"),
    ("SB-028", "壁挂式空调", "空调暖通", "星辰园", date(2021,4,22), 3500, "维修中", "李娜"),
    ("SB-029", "餐具消毒柜", "厨房电器", "星辰园", date(2022,8,30), 3200, "在用", "赵敏"),
    ("SB-030", "紫外线消毒灯", "厨房电器", "星辰园", date(2023,3,23), 900, "在用", "赵敏"),
    ("SB-031", "户外摇马组合", "户外设施", "星辰园", date(2021,10,9), 7600, "闲置", "陈静"),
    ("SB-032", "教学一体机", "教学设备", "星辰园", date(2024,9,1), 12800, "在用", "周燕"),
]
for i, (code, name, cat, park, dt, val, st, keeper) in enumerate(devices):
    r = 5 + i
    ws1.cell(row=r, column=2, value=code)
    ws1.cell(row=r, column=3, value=name)
    ws1.cell(row=r, column=4, value=cat)
    ws1.cell(row=r, column=5, value=park)
    c6 = ws1.cell(row=r, column=6, value=dt); c6.number_format = FMT_DATE
    c7 = ws1.cell(row=r, column=7, value=val); c7.number_format = FMT_CNY
    c8 = ws1.cell(row=r, column=8, value=f'=DATEDIF(F{r},TODAY(),"Y")')
    ws1.cell(row=r, column=9, value=st)
    ws1.cell(row=r, column=10, value=keeper)
    style_data_row(ws1, r, 2, LAST1, i)
    ws1.cell(row=r, column=6).alignment = Alignment(horizontal='center', vertical='center')
    for col in (7, 8):
        ws1.cell(row=r, column=col).alignment = Alignment(horizontal='right', vertical='center')
    ws1.cell(row=r, column=9).alignment = Alignment(horizontal='center', vertical='center')

T1 = 5 + len(devices)  # 37 合计行
ws1.cell(row=T1, column=2, value="合计")
c = ws1.cell(row=T1, column=7, value=f"=SUM(G5:G{T1-1})"); c.number_format = FMT_CNY
c = ws1.cell(row=T1, column=8, value=f"=ROUND(AVERAGE(H5:H{T1-1}),1)")
ws1.cell(row=T1, column=9, value=f'=COUNTA(I5:I{T1-1})&" 台"')
style_total_row(ws1, T1, 2, LAST1)
c.alignment = Alignment(horizontal='right', vertical='center')

note = ws1.cell(row=T1+2, column=2, value="注：使用状态分为「在用 / 维修中 / 闲置 / 待报废」四类，录入时请使用下拉列表选择；已用年限由公式自动计算。")
note.font = font_caption()

# 数据验证：使用状态下拉列表
dv1 = DataValidation(type="list", formula1='"在用,维修中,闲置,待报废"', allow_blank=False,
                     showErrorMessage=True, errorTitle="录入有误", error="请从下拉列表中选择使用状态")
ws1.add_data_validation(dv1); dv1.add(f"I5:I{T1-1}")
# 条件格式：维修中→红，待报废→琥珀
ws1.conditional_formatting.add(f"I5:I{T1-1}",
    CellIsRule(operator='equal', formula=['"维修中"'],
               fill=PatternFill('solid', fgColor='FDEDEC'), font=Font(name=base.FONT_NAME, color=ACC_NEG)))
ws1.conditional_formatting.add(f"I5:I{T1-1}",
    CellIsRule(operator='equal', formula=['"待报废"'],
               fill=PatternFill('solid', fgColor='FEF9E7'), font=Font(name=base.FONT_NAME, color=ACC_WARN)))

ws1.freeze_panes = "C5"
auto_fit_columns(ws1, header_row=4, data_start_row=5)
auto_fit_row_heights(ws1, header_row=4, data_start_row=5)

# ============================================================
# 工作表 2：维修记录
# ============================================================
ws2 = wb.create_sheet("维修记录")
H2 = ["报修单号", "设备编号", "设备名称", "园区", "故障类型", "报修日期", "完成日期", "维修状态", "维修费用(元)", "处理天数"]
LAST2 = 1 + len(H2)  # K
setup_sheet(ws2, title="设备维修记录（2026 年 6—10 月）", last_col=LAST2)
for c_, h in enumerate(H2, 2):
    ws2.cell(row=4, column=c_, value=h)
style_header_row(ws2, 4, 2, LAST2)

repairs = [
    # 单号, 设备编号, 故障类型, 报修日期, 完成日期, 状态, 费用
    ("WX-2606-01", "SB-002", "电路故障", date(2026,6,4),  date(2026,6,6),  "已完成", 260),
    ("WX-2606-02", "SB-006", "制冷异常", date(2026,6,11), date(2026,6,15), "已完成", 480),
    ("WX-2606-03", "SB-020", "耗材老化", date(2026,6,19), date(2026,6,22), "已完成", 150),
    ("WX-2607-01", "SB-010", "机械损坏", date(2026,7,3),  date(2026,7,10), "已完成", 1350),
    ("WX-2607-02", "SB-003", "软件故障", date(2026,7,9),  date(2026,7,11), "已完成", 100),
    ("WX-2607-03", "SB-022", "机械损坏", date(2026,7,16), date(2026,7,21), "已完成", 820),
    ("WX-2607-04", "SB-009", "管路漏水", date(2026,7,25), date(2026,7,28), "已完成", 350),
    ("WX-2608-01", "SB-014", "光学部件", date(2026,8,5),  date(2026,8,12), "已完成", 600),
    ("WX-2608-02", "SB-018", "制冷异常", date(2026,8,13), date(2026,8,16), "已完成", 420),
    ("WX-2608-03", "SB-005", "耗材老化", date(2026,8,22), date(2026,8,24), "已完成", 180),
    ("WX-2608-04", "SB-024", "软件故障", date(2026,8,28), date(2026,8,30), "已完成", 120),
    ("WX-2609-01", "SB-016", "进纸卡纸", date(2026,9,2),  date(2026,9,5),  "已完成", 220),
    ("WX-2609-02", "SB-028", "制冷异常", date(2026,9,8),  None,            "维修中", 300),
    ("WX-2609-03", "SB-029", "电路故障", date(2026,9,15), date(2026,9,17), "已完成", 240),
    ("WX-2609-04", "SB-031", "机械损坏", date(2026,9,21), date(2026,9,26), "已完成", 310),
    ("WX-2609-05", "SB-013", "触控失灵", date(2026,9,28), None,            "维修中", 260),
    ("WX-2610-01", "SB-004", "软件故障", date(2026,10,9), None,            "待派单", 0),
    ("WX-2610-02", "SB-006", "噪音异常", date(2026,10,14),date(2026,10,17),"已完成", 380),
    ("WX-2610-03", "SB-027", "电路故障", date(2026,10,20),None,            "待派单", 0),
    ("WX-2610-04", "SB-032", "触控失灵", date(2026,10,26),None,            "维修中", 200),
]
FAULT_MAP = {"电路故障","制冷异常","耗材老化","机械损坏","软件故障","管路漏水"}
for i, (no, dcode, fault, d1, d2, st, fee) in enumerate(repairs):
    r = 5 + i
    ws2.cell(row=r, column=2, value=no)
    ws2.cell(row=r, column=3, value=dcode)
    ws2.cell(row=r, column=4, value=f"=IFERROR(VLOOKUP($C{r},'设备台账'!$B$5:$C$36,2,FALSE),\"—\")")
    ws2.cell(row=r, column=5, value=f"=IFERROR(VLOOKUP($C{r},'设备台账'!$B$5:$E$36,4,FALSE),\"—\")")
    ws2.cell(row=r, column=6, value=fault)
    c7 = ws2.cell(row=r, column=7, value=d1); c7.number_format = FMT_DATE
    c8 = ws2.cell(row=r, column=8, value=d2 if d2 else None)
    if d2: c8.number_format = FMT_DATE
    ws2.cell(row=r, column=9, value=st)
    c10 = ws2.cell(row=r, column=10, value=fee); c10.number_format = FMT_CNY
    c11 = ws2.cell(row=r, column=11, value=f'=IF(H{r}="","",H{r}-G{r})')
    style_data_row(ws2, r, 2, LAST2, i)
    for col in (7, 8):
        ws2.cell(row=r, column=col).alignment = Alignment(horizontal='center', vertical='center')
    ws2.cell(row=r, column=9).alignment = Alignment(horizontal='center', vertical='center')
    ws2.cell(row=r, column=10).alignment = Alignment(horizontal='right', vertical='center')
    ws2.cell(row=r, column=11).alignment = Alignment(horizontal='center', vertical='center')

T2 = 5 + len(repairs)  # 25
ws2.cell(row=T2, column=2, value="合计")
c = ws2.cell(row=T2, column=10, value=f"=SUM(J5:J{T2-1})"); c.number_format = FMT_CNY
c = ws2.cell(row=T2, column=11, value=f'=COUNTA(B5:B{T2-1})&" 单"')
style_total_row(ws2, T2, 2, LAST2)
c.alignment = Alignment(horizontal='center', vertical='center')
ws2.cell(row=T2, column=10).alignment = Alignment(horizontal='right', vertical='center')

note2 = ws2.cell(row=T2+2, column=2, value="注：设备名称、园区由设备编号经 VLOOKUP 自动带出，无需手工录入；维修状态分为「待派单 / 维修中 / 已完成」。")
note2.font = font_caption()

dv2 = DataValidation(type="list", formula1='"待派单,维修中,已完成"', allow_blank=False,
                     showErrorMessage=True, errorTitle="录入有误", error="请从下拉列表中选择维修状态")
ws2.add_data_validation(dv2); dv2.add(f"I5:I{T2-1}")
dv3 = DataValidation(type="list", formula1='"电路故障,制冷异常,耗材老化,机械损坏,软件故障,管路漏水,触控失灵,光学部件,进纸卡纸,噪音异常"',
                     allow_blank=False, showErrorMessage=True, errorTitle="录入有误", error="请选择标准故障类型，新类型请联系管理员维护")
ws2.add_data_validation(dv3); dv3.add(f"F5:F{T2-1}")
ws2.conditional_formatting.add(f"I5:I{T2-1}",
    CellIsRule(operator='equal', formula=['"待派单"'],
               fill=PatternFill('solid', fgColor='FDEDEC'), font=Font(name=base.FONT_NAME, color=ACC_NEG)))
ws2.conditional_formatting.add(f"I5:I{T2-1}",
    CellIsRule(operator='equal', formula=['"维修中"'],
               fill=PatternFill('solid', fgColor='FEF9E7'), font=Font(name=base.FONT_NAME, color=ACC_WARN)))
ws2.conditional_formatting.add(f"I5:I{T2-1}",
    CellIsRule(operator='equal', formula=['"已完成"'],
               fill=PatternFill('solid', fgColor='E8F5E9'), font=Font(name=base.FONT_NAME, color=ACC_POS)))

ws2.freeze_panes = "C5"
auto_fit_columns(ws2, header_row=4, data_start_row=5)
auto_fit_row_heights(ws2, header_row=4, data_start_row=5)

# ============================================================
# 工作表 3：月度统计（公式 + 透视统计区）
# ============================================================
ws3 = wb.create_sheet("月度统计")
setup_sheet(ws3, title="园务设备月度统计（截至 2026 年 10 月）", last_col=8)

def subheader(ws, row, text, last_col=8):
    cell = ws.cell(row=row, column=2, value=text)
    cell.font = Font(name=base.FONT_NAME, size=12, bold=True, color=PRIMARY)
    ws.row_dimensions[row].height = 24

def write_block_header(ws, row, headers, col_start=2):
    for c_, h in enumerate(headers, col_start):
        ws.cell(row=row, column=c_, value=h)
    style_header_row(ws, row, col_start, col_start + len(headers) - 1)

def style_block_rows(ws, start, end, col_start, col_end, fmts=None):
    for i, r in enumerate(range(start, end + 1)):
        style_data_row(ws, r, col_start, col_end, i)
        for c_, fmt in (fmts or {}).items():
            cell = ws.cell(row=r, column=c_)
            cell.number_format = fmt
            cell.alignment = Alignment(horizontal='right' if fmt in (FMT_INT, FMT_CNY, FMT_PCT) else 'center', vertical='center')

PARKS = ["阳光园", "彩虹园", "星辰园"]
STATUSES = ["在用", "维修中", "闲置", "待报废"]
FAULTS = ["机械损坏", "制冷异常", "电路故障", "软件故障", "耗材老化", "触控失灵"]
OTHER_FAULTS = ["管路漏水", "光学部件", "进纸卡纸", "噪音异常"]
MONTHS = [("2026-06", 6), ("2026-07", 7), ("2026-08", 8), ("2026-09", 9), ("2026-10", 10)]

# —— 区块 1：各园区设备概况（行 4 表头）——
subheader(ws3, 3, "一、各园区设备概况")
write_block_header(ws3, 4, ["园区", "设备总数", "维修中", "在用", "闲置", "待报废", "维修中占比"])
TD = "'设备台账'"
r0 = 5
for i, p in enumerate(PARKS):
    r = r0 + i
    ws3.cell(row=r, column=2, value=p)
    ws3.cell(row=r, column=3, value=f"=COUNTIF({TD}!$E$5:$E$36,$B{r})")
    ws3.cell(row=r, column=4, value=f'=COUNTIFS({TD}!$E$5:$E$36,$B{r},{TD}!$I$5:$I$36,"维修中")')
    ws3.cell(row=r, column=5, value=f'=COUNTIFS({TD}!$E$5:$E$36,$B{r},{TD}!$I$5:$I$36,"在用")')
    ws3.cell(row=r, column=6, value=f'=COUNTIFS({TD}!$E$5:$E$36,$B{r},{TD}!$I$5:$I$36,"闲置")')
    ws3.cell(row=r, column=7, value=f'=COUNTIFS({TD}!$E$5:$E$36,$B{r},{TD}!$I$5:$I$36,"待报废")')
    ws3.cell(row=r, column=8, value=f"=IFERROR(D{r}/C{r},0)")
r1 = r0 + len(PARKS) - 1  # 7
rTot = r1 + 1  # 8
ws3.cell(row=rTot, column=2, value="合计")
for col in range(3, 8):
    L = get_column_letter(col)
    ws3.cell(row=rTot, column=col, value=f"=SUM({L}{r0}:{L}{r1})")
ws3.cell(row=rTot, column=8, value=f"=IFERROR(D{rTot}/C{rTot},0)")
style_block_rows(ws3, r0, r1, 2, 8, {3: FMT_INT, 4: FMT_INT, 5: FMT_INT, 6: FMT_INT, 7: FMT_INT, 8: FMT_PCT})
style_total_row(ws3, rTot, 2, 8)
for col in range(3, 8):
    ws3.cell(row=rTot, column=col).alignment = Alignment(horizontal='right', vertical='center')
ws3.cell(row=rTot, column=8).number_format = FMT_PCT
ws3.cell(row=rTot, column=8).alignment = Alignment(horizontal='right', vertical='center')

# —— 区块 2：设备使用状态分布（行 12 表头）——
rowA = rTot + 3  # 11 小节标题
subheader(ws3, rowA, "二、设备使用状态分布")
hB = rowA + 1    # 12 表头
write_block_header(ws3, hB, ["使用状态", "设备数量", "占比"])
b0 = hB + 1      # 13
for i, st in enumerate(STATUSES):
    r = b0 + i
    ws3.cell(row=r, column=2, value=st)
    ws3.cell(row=r, column=3, value=f"=COUNTIF({TD}!$I$5:$I$36,$B{r})")
    ws3.cell(row=r, column=4, value=f"=IFERROR(C{r}/COUNTA({TD}!$B$5:$B$36),0)")
b1 = b0 + len(STATUSES) - 1  # 16
bTot = b1 + 1  # 17
ws3.cell(row=bTot, column=2, value="合计")
ws3.cell(row=bTot, column=3, value=f"=SUM(C{b0}:C{b1})")
ws3.cell(row=bTot, column=4, value=f"=SUM(D{b0}:D{b1})")
style_block_rows(ws3, b0, b1, 2, 4, {3: FMT_INT, 4: FMT_PCT})
style_total_row(ws3, bTot, 2, 4)
ws3.cell(row=bTot, column=3).alignment = Alignment(horizontal='right', vertical='center')
ws3.cell(row=bTot, column=4).number_format = FMT_PCT
ws3.cell(row=bTot, column=4).alignment = Alignment(horizontal='right', vertical='center')

# —— 区块 3：故障类型分布（行 21 表头）——
rowC = bTot + 3  # 20 小节标题
subheader(ws3, rowC, "三、故障类型分布（2026 年 6—10 月）")
hC = rowC + 1
write_block_header(ws3, hC, ["故障类型", "报修数量", "维修费用(元)", "费用占比"])
c0 = hC + 1
WR = "'维修记录'"
for i, f_ in enumerate(FAULTS):
    r = c0 + i
    ws3.cell(row=r, column=2, value=f_)
    ws3.cell(row=r, column=3, value=f"=COUNTIF({WR}!$F$5:$F$24,$B{r})")
    ws3.cell(row=r, column=4, value=f"=SUMIF({WR}!$F$5:$F$24,$B{r},{WR}!$J$5:$J$24)")
# 「其他」行：汇总低频故障类型，保证口径完整覆盖全部报修单
rO = c0 + len(FAULTS)
ws3.cell(row=rO, column=2, value="其他")
ws3.cell(row=rO, column=3, value="=" + "+".join(f'COUNTIF({WR}!$F$5:$F$24,"{f}")' for f in OTHER_FAULTS))
ws3.cell(row=rO, column=4, value="=" + "+".join(f'SUMIF({WR}!$F$5:$F$24,"{f}",{WR}!$J$5:$J$24)' for f in OTHER_FAULTS))
c1 = rO
cTot = c1 + 1
ws3.cell(row=cTot, column=2, value="合计")
ws3.cell(row=cTot, column=3, value=f"=SUM(C{c0}:C{c1})")
ws3.cell(row=cTot, column=4, value=f"=SUM(D{c0}:D{c1})")
ws3.cell(row=cTot, column=5, value=f"=SUM(E{c0}:E{c1})")
style_block_rows(ws3, c0, c1, 2, 5, {3: FMT_INT, 4: FMT_CNY, 5: FMT_PCT})
for r in range(c0, c1 + 1):  # 费用占比 = 本类费用/总费用
    ws3.cell(row=r, column=5, value=f"=IFERROR(D{r}/$D${cTot},0)")
    ws3.cell(row=r, column=5).number_format = FMT_PCT
style_total_row(ws3, cTot, 2, 5)
for col in (3, 4, 5):
    ws3.cell(row=cTot, column=col).alignment = Alignment(horizontal='right', vertical='center')
ws3.cell(row=cTot, column=3).number_format = FMT_INT
ws3.cell(row=cTot, column=4).number_format = FMT_CNY
ws3.cell(row=cTot, column=5).number_format = FMT_PCT

# —— 区块 4：月度报修与维修趋势（行 32 表头）——
rowD = cTot + 3
subheader(ws3, rowD, "四、月度报修与维修趋势")
hD = rowD + 1
write_block_header(ws3, hD, ["月份", "报修数量", "已完成", "完成率", "维修费用(元)"])
d0 = hD + 1
for i, (label, m) in enumerate(MONTHS):
    r = d0 + i
    m_start = date(2026, m, 1)
    m_end = date(2026, m + 1, 1) if m < 12 else date(2027, 1, 1)
    ws3.cell(row=r, column=2, value=label)
    ws3.cell(row=r, column=3, value=f'=COUNTIFS({WR}!$G$5:$G$24,">="&DATE(2026,{m},1),{WR}!$G$5:$G$24,"<"&DATE(2026,{m+1},1))')
    ws3.cell(row=r, column=4, value=f'=COUNTIFS({WR}!$G$5:$G$24,">="&DATE(2026,{m},1),{WR}!$G$5:$G$24,"<"&DATE(2026,{m+1},1),{WR}!$I$5:$I$24,"已完成")')
    ws3.cell(row=r, column=5, value=f"=IFERROR(D{r}/C{r},0)")
    ws3.cell(row=r, column=6, value=f'=SUMIFS({WR}!$J$5:$J$24,{WR}!$G$5:$G$24,">="&DATE(2026,{m},1),{WR}!$G$5:$G$24,"<"&DATE(2026,{m+1},1))')
d1 = d0 + len(MONTHS) - 1
dTot = d1 + 1
ws3.cell(row=dTot, column=2, value="合计")
for col, fmt in ((3, FMT_INT), (4, FMT_INT), (6, FMT_CNY)):
    L = get_column_letter(col)
    cc = ws3.cell(row=dTot, column=col, value=f"=SUM({L}{d0}:{L}{d1})")
    cc.number_format = fmt
ws3.cell(row=dTot, column=5, value=f"=IFERROR(D{dTot}/C{dTot},0)")
style_block_rows(ws3, d0, d1, 2, 6, {3: FMT_INT, 4: FMT_INT, 5: FMT_PCT, 6: FMT_CNY})
style_total_row(ws3, dTot, 2, 6)
for col in range(3, 7):
    ws3.cell(row=dTot, column=col).alignment = Alignment(horizontal='right', vertical='center')
ws3.cell(row=dTot, column=3).number_format = FMT_INT
ws3.cell(row=dTot, column=4).number_format = FMT_INT
ws3.cell(row=dTot, column=5).number_format = FMT_PCT
ws3.cell(row=dTot, column=6).number_format = FMT_CNY

note3 = ws3.cell(row=dTot + 2, column=2, value="注：本表全部数值由 COUNTIF / COUNTIFS / SUMIF / SUMIFS 公式对「设备台账」「维修记录」实时统计生成；数据透视分析见「数据透视」工作表。")
note3.font = font_caption()

ws3.column_dimensions['A'].width = 3
for col, w in {'B': 12, 'C': 13, 'D': 12, 'E': 14, 'F': 14, 'G': 12, 'H': 12}.items():
    ws3.column_dimensions[col].width = w

# ============================================================
# 工作表 4：管理看板（KPI + 图表）
# ============================================================
ws4 = wb.create_sheet("管理看板")
setup_sheet(ws4, title="园务设备管理看板（数据截至 2026 年 10 月）", last_col=13)

kpis = [
    ("设备总数(台)", f"='月度统计'!C{rTot}", FMT_INT),
    ("维修中(台)", f"='月度统计'!D{rTot}", FMT_INT),
    ("累计报修(单)", f"='月度统计'!C{dTot}", FMT_INT),
    ("维修完成率", f"='月度统计'!E{dTot}", FMT_PCT),
    ("维修费用合计", f"='月度统计'!F{dTot}", FMT_CNY),
]
for i, (label, formula, fmt) in enumerate(kpis):
    col = 2 + i * 2  # B, D, F, H, J
    lc = ws4.cell(row=4, column=col, value=label)
    lc.font = font_kpi_label(); lc.alignment = Alignment(horizontal='left', vertical='center')
    vc = ws4.cell(row=5, column=col, value=formula)
    vc.font = font_kpi(); vc.number_format = fmt
    vc.alignment = Alignment(horizontal='left', vertical='center')
ws4.row_dimensions[5].height = 30

# 图表 1：各园区设备数量（柱形，总数 + 维修中）
bar = create_bar_chart()
setup_chart_titles(bar, title="各园区设备数量对比", y_title="数量(台)", x_title="园区")
bar.add_data(Reference(ws3, min_col=3, max_col=4, min_row=4, max_row=r1), titles_from_data=True)
bar.set_categories(Reference(ws3, min_col=2, min_row=r0, max_row=r1))
apply_chart_colors(bar)
bar.width, bar.height = 16, 9.5
ws4.add_chart(bar, "B8")

# 图表 2：故障类型分布（饼图）
pie = create_pie_chart()
setup_chart_titles(pie, title="故障类型分布")
pie.add_data(Reference(ws3, min_col=3, min_row=hC, max_row=c1), titles_from_data=True)
pie.set_categories(Reference(ws3, min_col=2, min_row=c0, max_row=c1))
from openpyxl.chart.label import DataLabelList
pie.dataLabels = DataLabelList()
pie.dataLabels.showCatName = True
pie.dataLabels.showPercent = True
pie.dataLabels.showVal = False
pie.dataLabels.showLeaderLines = True
pie.width, pie.height = 14, 9.5
ws4.add_chart(pie, "H8")

# 图表 3：月度报修与完成趋势（折线）
line = create_line_chart()
setup_chart_titles(line, title="月度报修与维修完成趋势", y_title="数量(单)", x_title="月份")
line.add_data(Reference(ws3, min_col=3, max_col=4, min_row=hD, max_row=d1), titles_from_data=True)
line.set_categories(Reference(ws3, min_col=2, min_row=d0, max_row=d1))
for s in line.series:
    s.smooth = False
apply_chart_colors(line)
line.width, line.height = 16, 9.5
ws4.add_chart(line, "B27")

# 图表 4：月度维修费用（柱形）
bar2 = create_bar_chart()
setup_chart_titles(bar2, title="月度维修费用", y_title="费用(元)", x_title="月份")
bar2.add_data(Reference(ws3, min_col=6, max_col=6, min_row=hD, max_row=d1), titles_from_data=True)
bar2.set_categories(Reference(ws3, min_col=2, min_row=d0, max_row=d1))
apply_chart_colors(bar2)
bar2.width, bar2.height = 14, 9.5
ws4.add_chart(bar2, "H27")

note4 = ws4.cell(row=46, column=2, value="注：看板指标与图表均引用「月度统计」工作表，源数据更新后自动刷新。")
note4.font = font_caption()
ws4.column_dimensions['A'].width = 3
for col in range(2, 14):
    ws4.column_dimensions[get_column_letter(col)].width = 11

# 打印设置
for ws in (ws1, ws2):
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = '4:4'

wb.properties.creator = "Z.ai"
wb.save(OUT)
print("saved:", OUT)
