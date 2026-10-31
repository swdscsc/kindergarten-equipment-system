# -*- coding: utf-8 -*-
"""园务设备管理系统 — 功能测试用例（pytest + Flask test client）。"""
import io

from openpyxl import load_workbook


# ─────────────── 管理看板 ───────────────
class TestDashboard:
    def test_dashboard_200_with_kpi(self, client):
        r = client.get("/")
        assert r.status_code == 200
        html = r.get_data(as_text=True)
        assert "管理看板" in html and "32" in html  # 种子数据 32 台设备

    def test_api_stats_aggregates(self, client):
        data = client.get("/api/stats").get_json()
        assert sum(data["parkTotal"]) == 32
        assert sum(data["parkFixing"]) == 4
        assert sum(data["reported"]) == 20
        assert sum(data["finished"]) == 15
        assert sum(f["value"] for f in data["faults"]) == 20


# ─────────────── 设备台账 ───────────────
class TestDevices:
    def test_list_contains_seed_devices(self, client):
        html = client.get("/devices").get_data(as_text=True)
        assert "SB-001" in html and "SB-032" in html and "教学一体机" in html

    def test_filter_by_park(self, client):
        html = client.get("/devices?park=星辰园").get_data(as_text=True)
        assert "SB-024" in html and "SB-001" not in html

    def test_filter_by_status(self, client):
        html = client.get("/devices?status=维修中").get_data(as_text=True)
        assert "SB-004" in html and "SB-003" not in html

    def test_search_by_keyword(self, client):
        html = client.get("/devices?q=消毒柜").get_data(as_text=True)
        assert "SB-008" in html and "SB-001" not in html

    def test_create_device_success(self, client):
        r = client.post("/devices/new", data={
            "code": "SB-101", "name": "空气消毒机", "category": "厨房电器",
            "park": "阳光园", "purchase_date": "2026-10-15", "value": "3000",
            "status": "在用", "keeper": "李娜"}, follow_redirects=True)
        html = r.get_data(as_text=True)
        assert "设备已新增：SB-101" in html and "SB-101" in html

    def test_create_device_duplicate_code_rejected(self, client):
        r = client.post("/devices/new", data={
            "code": "SB-001", "name": "重复编号", "category": "教学设备",
            "park": "阳光园", "purchase_date": "2026-01-01", "value": "100",
            "status": "在用", "keeper": "张莉"}, follow_redirects=True)
        assert "设备编号已存在" in r.get_data(as_text=True)

    def test_create_device_invalid_date_rejected(self, client):
        r = client.post("/devices/new", data={
            "code": "SB-102", "name": "日期错误", "category": "教学设备",
            "park": "阳光园", "purchase_date": "2026/01/01", "value": "100",
            "status": "在用", "keeper": "张莉"}, follow_redirects=True)
        assert "采购日期格式" in r.get_data(as_text=True)

    def test_edit_device_updates_record(self, client):
        r = client.post("/devices/1/edit", data={
            "code": "SB-001", "name": "教学一体机", "category": "教学设备",
            "park": "阳光园", "purchase_date": "2022-03-15", "value": "13800",
            "status": "闲置", "keeper": "张莉"}, follow_redirects=True)
        html = r.get_data(as_text=True)
        assert "设备已更新：SB-001" in html and "闲置" in html

    def test_delete_device_without_repairs(self, client):
        r = client.post("/devices/12/delete", follow_redirects=True)  # SB-012 无维修记录
        assert "设备已删除" in r.get_data(as_text=True)

    def test_delete_device_with_repairs_blocked(self, client):
        r = client.post("/devices/2/delete", follow_redirects=True)  # SB-002 有维修单
        html = r.get_data(as_text=True)
        assert "禁止删除" in html and "SB-002" in html


# ─────────────── 维修记录 ───────────────
class TestRepairs:
    def test_list_contains_seed_orders(self, client):
        html = client.get("/repairs").get_data(as_text=True)
        assert "WX-2610-04" in html and "触控失灵" in html

    def test_create_repair_auto_order_no_and_status_sync(self, client):
        r = client.post("/repairs/new", data={
            "device_code": "SB-010", "fault_type": "机械损坏",
            "report_date": "2026-10-28"}, follow_redirects=True)
        html = r.get_data(as_text=True)
        assert "WX-2610-05" in html          # 自动编号：当月已有 2 单，顺序 +1
        assert "报修单已登记" in html
        # 设备状态联动：在用 → 维修中
        dev_html = client.get("/devices?q=SB-010").get_data(as_text=True)
        assert 'tag-维修中' in dev_html

    def test_create_repair_unknown_device_rejected(self, client):
        r = client.post("/repairs/new", data={
            "device_code": "SB-999", "fault_type": "机械损坏",
            "report_date": "2026-10-28"}, follow_redirects=True)
        assert "设备编号不存在" in r.get_data(as_text=True)

    def test_create_repair_invalid_fault_type_rejected(self, client):
        r = client.post("/repairs/new", data={
            "device_code": "SB-010", "fault_type": "不明原因",
            "report_date": "2026-10-28"}, follow_redirects=True)
        assert "标准故障类型" in r.get_data(as_text=True)

    def test_update_to_done_requires_finish_date(self, client):
        # WX-2610-01（id=17）当前待派单 → 缺完成日期应被拒绝
        r = client.post("/repairs/17/update", data={"status": "已完成", "cost": "100"},
                        follow_redirects=True)
        assert "完成日期" in r.get_data(as_text=True)

    def test_update_to_done_restores_device_status(self, client):
        # WX-2609-02（id=13）维修中 → 已完成，设备 SB-028 恢复在用
        r = client.post("/repairs/13/update", data={
            "status": "已完成", "finish_date": "2026-10-30", "cost": "360"},
            follow_redirects=True)
        html = r.get_data(as_text=True)
        assert "已更新为「已完成」" in html
        dev_html = client.get("/devices?q=SB-028").get_data(as_text=True)
        assert 'tag-在用' in dev_html

    def test_repair_filter_by_status(self, client):
        html = client.get("/repairs?status=待派单").get_data(as_text=True)
        assert "WX-2610-01" in html and "WX-2610-02" not in html


# ─────────────── Excel 导出 ───────────────
class TestExport:
    def test_export_returns_valid_workbook(self, client, app):
        r = client.get("/export.xlsx")
        assert r.status_code == 200
        assert r.content_type.startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        wb = load_workbook(io.BytesIO(r.data))
        assert wb.sheetnames == ["设备台账", "维修记录", "月度统计", "管理看板"]
        ws = wb["设备台账"]
        assert ws["B5"].value == "SB-001"
        assert ws.max_row >= 36  # 表头 4 行 + 32 台设备
        ws2 = wb["维修记录"]
        assert ws2["B5"].value == "WX-2606-01"
