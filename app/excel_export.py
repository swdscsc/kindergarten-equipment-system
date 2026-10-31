# -*- coding: utf-8 -*-
"""Excel 导出：从数据库实时生成《园务设备管理系统》工作簿（四表 + 透视统计区）。"""
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "微软雅黑"
PRIMARY, ACC_POS, ACC_NEG, ACC_WARN = "1B2A4A", "1B7D46", "C0392B", "D4820A"
N900, N600, N100, N0 = "37352F", "8C8A84", "F7F7F5", "FFFFFF"
FMT_CNY, FMT_DATE, FMT_PCT = "¥#,##0.00", "yyyy-mm-dd", "0.0%"


def _style_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor=PRIMARY)
        cell.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def _style_rows(ws, r1, r2, c1, c2):
    for i, r in enumerate(range(r1, r2 + 1)):
        fill = PatternFill("solid", fgColor=(N0 if i % 2 == 0 else N100))
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill
            cell.font = Font(name=FONT, size=11, color=N900)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 22


def _title(ws, text, last_col):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.row_dimensions[1].height = 15
    ws.row_dimensions[2].height = 32
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col)
    c = ws.cell(row=2, column=2, value=text)
    c.font = Font(name=FONT, size=16, bold=True, color=PRIMARY)
    c.alignment = Alignment(horizontal="left", vertical="center")


def build_workbook(conn):
    wb = Workbook()
    devices = conn.execute("SELECT * FROM devices ORDER BY code").fetchall()
    repairs = conn.execute(
        "SELECT r.order_no, d.code, d.name, d.park, r.fault_type, r.report_date,"
        " r.finish_date, r.status, r.cost FROM repairs r JOIN devices d ON d.id=r.device_id"
        " ORDER BY r.report_date, r.order_no").fetchall()

    # ── 设备台账 ──
    ws = wb.active
    ws.title = "设备台账"
    H = ["设备编号", "设备名称", "类别", "园区", "采购日期", "采购原值(元)", "使用状态", "保管人"]
    _title(ws, "幼儿园集团设备台账（系统导出）", 1 + len(H))
    for c, h in enumerate(H, 2):
        ws.cell(row=4, column=c, value=h)
    _style_header(ws, 4, 2, 1 + len(H))
    for i, d in enumerate(devices):
        r = 5 + i
        vals = [d["code"], d["name"], d["category"], d["park"], d["purchase_date"],
                d["value"], d["status"], d["keeper"]]
        for c, v in enumerate(vals, 2):
            ws.cell(row=r, column=c, value=v)
        ws.cell(row=r, column=6).number_format = FMT_CNY
        ws.cell(row=r, column=6).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=r, column=7).alignment = Alignment(horizontal="center", vertical="center")
    _style_rows(ws, 5, 4 + len(devices), 2, 1 + len(H))
    dv = DataValidation(type="list", formula1='"在用,维修中,闲置,待报废"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"H5:H{4 + len(devices)}")
    ws.conditional_formatting.add(
        f"H5:H{4 + len(devices)}",
        CellIsRule(operator="equal", formula=['"维修中"'],
                   fill=PatternFill("solid", fgColor="FDEDEC"), font=Font(color=ACC_NEG)))
    ws.conditional_formatting.add(
        f"H5:H{4 + len(devices)}",
        CellIsRule(operator="equal", formula=['"待报废"'],
                   fill=PatternFill("solid", fgColor="FEF9E7"), font=Font(color=ACC_WARN)))
    for col, w in {"B": 11, "C": 15, "D": 11, "E": 10, "F": 13, "G": 14, "H": 10, "I": 9}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C5"

    # ── 维修记录 ──
    ws2 = wb.create_sheet("维修记录")
    H2 = ["报修单号", "设备编号", "设备名称", "园区", "故障类型", "报修日期", "完成日期", "维修状态", "维修费用(元)"]
    _title(ws2, "设备维修记录（系统导出）", 1 + len(H2))
    for c, h in enumerate(H2, 2):
        ws2.cell(row=4, column=c, value=h)
    _style_header(ws2, 4, 2, 1 + len(H2))
    for i, r0 in enumerate(repairs):
        r = 5 + i
        vals = [r0["order_no"], r0["code"], r0["name"], r0["park"], r0["fault_type"],
                r0["report_date"], r0["finish_date"], r0["status"], r0["cost"]]
        for c, v in enumerate(vals, 2):
            ws2.cell(row=r, column=c, value=v)
        ws2.cell(row=r, column=10).number_format = FMT_CNY
        ws2.cell(row=r, column=10).alignment = Alignment(horizontal="right", vertical="center")
    _style_rows(ws2, 5, 4 + len(repairs), 2, 1 + len(H2))
    dv2 = DataValidation(type="list", formula1='"待派单,维修中,已完成"', allow_blank=False)
    ws2.add_data_validation(dv2)
    dv2.add(f"I5:I{4 + len(repairs)}")
    for col, w in {"B": 13, "C": 11, "D": 15, "E": 10, "F": 11, "G": 13, "H": 13, "I": 10, "J": 13}.items():
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "C5"

    # ── 月度统计（公式实时汇总）──
    ws3 = wb.create_sheet("月度统计")
    _title(ws3, "园务设备月度统计（公式实时汇总）", 7)
    n_dev, n_rep = len(devices), len(repairs)
    dev_end, rep_end = 4 + n_dev, 4 + n_rep
    TD, WR = "'设备台账'", "'维修记录'"

    def block(ws_, row, headers, formulas, fmts):
        for c, h in enumerate(headers, 2):
            ws_.cell(row=row, column=c, value=h)
        _style_header(ws_, row, 2, 1 + len(headers))
        for i, frow in enumerate(formulas):
            r = row + 1 + i
            for c, v in enumerate(frow, 2):
                ws_.cell(row=r, column=c, value=v)
        _style_rows(ws_, row + 1, row + len(formulas), 2, 1 + len(headers))
        for r in range(row + 1, row + 1 + len(formulas)):
            for c, fmt in fmts.items():
                cell = ws_.cell(row=r, column=c)
                cell.number_format = fmt
                cell.alignment = Alignment(horizontal="right", vertical="center")

    parks = ["阳光园", "彩虹园", "星辰园"]
    stat = [["园区", "设备总数", "维修中", "在用", "闲置", "待报废", "维修中占比"]]
    for i, p in enumerate(parks):
        r = 5 + i
        stat.append([p,
                     f"=COUNTIF({TD}!$E$5:$E${dev_end},$B{r})",
                     f'=COUNTIFS({TD}!$E$5:$E${dev_end},$B{r},{TD}!$H$5:$H${dev_end},"维修中")',
                     f'=COUNTIFS({TD}!$E$5:$E${dev_end},$B{r},{TD}!$H$5:$H${dev_end},"在用")',
                     f'=COUNTIFS({TD}!$E$5:$E${dev_end},$B{r},{TD}!$H$5:$H${dev_end},"闲置")',
                     f'=COUNTIFS({TD}!$E$5:$E${dev_end},$B{r},{TD}!$H$5:$H${dev_end},"待报废")',
                     f"=IFERROR(D{r}/C{r},0)"])
    ws3.cell(row=3, column=2, value="一、各园区设备概况").font = Font(name=FONT, size=12, bold=True, color=PRIMARY)
    block(ws3, 4, stat[0], stat[1:], {3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: FMT_PCT})

    row = 11
    ws3.cell(row=row, column=2, value="二、故障类型分布").font = Font(name=FONT, size=12, bold=True, color=PRIMARY)
    faults = conn.execute("SELECT fault_type f, COUNT(*) c, SUM(cost) s FROM repairs GROUP BY fault_type ORDER BY c DESC, f").fetchall()
    fh = ["故障类型", "报修数量", "维修费用(元)"]
    for c, h in enumerate(fh, 2):
        ws3.cell(row=row + 1, column=c, value=h)
    _style_header(ws3, row + 1, 2, 1 + len(fh))
    for i, fr in enumerate(faults):
        r = row + 2 + i
        ws3.cell(row=r, column=2, value=fr["f"])
        ws3.cell(row=r, column=3, value=fr["c"])
        ws3.cell(row=r, column=4, value=fr["s"])
    _style_rows(ws3, row + 2, row + 1 + len(faults), 2, 4)
    for r in range(row + 2, row + 2 + len(faults)):
        ws3.cell(row=r, column=3).alignment = Alignment(horizontal="right", vertical="center")
        c4 = ws3.cell(row=r, column=4)
        c4.number_format = FMT_CNY
        c4.alignment = Alignment(horizontal="right", vertical="center")

    row = row + 2 + len(faults) + 2
    ws3.cell(row=row, column=2, value="三、月度报修与维修趋势").font = Font(name=FONT, size=12, bold=True, color=PRIMARY)
    months = ["2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
    mh = ["月份", "报修数量", "已完成", "完成率", "维修费用(元)"]
    for c, h in enumerate(mh, 2):
        ws3.cell(row=row + 1, column=c, value=h)
    _style_header(ws3, row + 1, 2, 1 + len(mh))
    for i, m in enumerate(months):
        r = row + 2 + i
        mm = int(m[5:7])
        ws3.cell(row=r, column=2, value=m)
        ws3.cell(row=r, column=3, value=f'=COUNTIFS({WR}!$G$5:$G${rep_end},">="&DATE({m[:4]},{mm},1),{WR}!$G$5:$G${rep_end},"<"&DATE({m[:4]},{mm + 1},1))')
        ws3.cell(row=r, column=4, value=f'=COUNTIFS({WR}!$G$5:$G${rep_end},">="&DATE({m[:4]},{mm},1),{WR}!$G$5:$G${rep_end},"<"&DATE({m[:4]},{mm + 1},1),{WR}!$I$5:$I${rep_end},"已完成")')
        ws3.cell(row=r, column=5, value=f"=IFERROR(D{r}/C{r},0)")
        ws3.cell(row=r, column=6, value=f'=SUMIFS({WR}!$J$5:$J${rep_end},{WR}!$G$5:$G${rep_end},">="&DATE({m[:4]},{mm},1),{WR}!$G$5:$G${rep_end},"<"&DATE({m[:4]},{mm + 1},1))')
    _style_rows(ws3, row + 2, row + 1 + len(months), 2, 6)
    for r in range(row + 2, row + 2 + len(months)):
        for c, fmt in ((3, "#,##0"), (4, "#,##0"), (5, FMT_PCT), (6, FMT_CNY)):
            cell = ws3.cell(row=r, column=c)
            cell.number_format = fmt
            cell.alignment = Alignment(horizontal="right", vertical="center")
    for col, w in {"B": 13, "C": 12, "D": 12, "E": 13, "F": 14, "G": 13, "H": 12}.items():
        ws3.column_dimensions[col].width = w

    # ── 管理看板（KPI）──
    ws4 = wb.create_sheet("管理看板")
    _title(ws4, "园务设备管理看板（KPI 汇总）", 10)
    kpis = [("设备总数(台)", f"=COUNTA({TD}!$B$5:$B${dev_end})", "#,##0"),
            ("在用占比", f'=COUNTIF({TD}!$H$5:$H${dev_end},"在用")/COUNTA({TD}!$B$5:$B${dev_end})', FMT_PCT),
            ("维修中(台)", f'=COUNTIF({TD}!$H$5:$H${dev_end},"维修中")', "#,##0"),
            ("累计报修(单)", f"=COUNTA({WR}!$B$5:$B${rep_end})", "#,##0"),
            ("维修完成率", f'=COUNTIF({WR}!$I$5:$I${rep_end},"已完成")/COUNTA({WR}!$B$5:$B${rep_end})', FMT_PCT),
            ("维修费用合计", f"=SUM({WR}!$J$5:$J${rep_end})", FMT_CNY)]
    for i, (label, formula, fmt) in enumerate(kpis):
        col = 2 + i
        lc = ws4.cell(row=4, column=col, value=label)
        lc.font = Font(name=FONT, size=9, color=N600)
        vc = ws4.cell(row=5, column=col, value=formula)
        vc.font = Font(name=FONT, size=20, bold=True, color=PRIMARY)
        vc.number_format = fmt
    ws4.row_dimensions[5].height = 30
    for col in range(2, 8):
        ws4.column_dimensions[get_column_letter(col)].width = 14

    wb.properties.creator = "Z.ai"
    return wb
