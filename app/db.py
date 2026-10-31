# -*- coding: utf-8 -*-
"""数据库访问层：SQLite 连接、建表、初始化与示例数据。"""
import os
import sqlite3

from flask import current_app, g

DB_NAME = "kindergarten.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS devices (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    code          TEXT NOT NULL UNIQUE,
    name          TEXT NOT NULL,
    category      TEXT NOT NULL,
    park          TEXT NOT NULL,
    purchase_date TEXT NOT NULL,
    value         REAL NOT NULL,
    status        TEXT NOT NULL DEFAULT '在用',
    keeper        TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS repairs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    order_no    TEXT NOT NULL UNIQUE,
    device_id   INTEGER NOT NULL REFERENCES devices(id),
    fault_type  TEXT NOT NULL,
    report_date TEXT NOT NULL,
    finish_date TEXT,
    status      TEXT NOT NULL DEFAULT '待派单',
    cost        REAL NOT NULL DEFAULT 0
);
"""

PARKS = ["阳光园", "彩虹园", "星辰园"]
DEVICE_STATUS = ["在用", "维修中", "闲置", "待报废"]
REPAIR_STATUS = ["待派单", "维修中", "已完成"]
FAULT_TYPES = ["机械损坏", "制冷异常", "电路故障", "软件故障", "耗材老化",
               "触控失灵", "管路漏水", "光学部件", "进纸卡纸", "噪音异常"]
CATEGORIES = ["教学设备", "办公设备", "厨房电器", "空调暖通", "户外设施"]


def db_path(app):
    return os.path.join(app.instance_path, DB_NAME)


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(current_app.config["DB_PATH"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def close_db(e=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db(app, seed=True):
    os.makedirs(app.instance_path, exist_ok=True)
    conn = sqlite3.connect(db_path(app))
    conn.executescript(SCHEMA)
    if seed:
        count = conn.execute("SELECT COUNT(*) FROM devices").fetchone()[0]
        if count == 0:
            conn.executemany(
                "INSERT INTO devices (code,name,category,park,purchase_date,value,status,keeper)"
                " VALUES (?,?,?,?,?,?,?,?)", SEED_DEVICES)
            conn.executemany(
                "INSERT INTO repairs (order_no,device_id,fault_type,report_date,finish_date,status,cost)"
                " VALUES (?,?,?,?,?,?,?)",
                [(o, _did(code), f, r, fi, s, c)
                 for o, code, f, r, fi, s, c in SEED_REPAIRS])
    conn.commit()
    conn.close()


def _did(code):
    """seed 阶段的临时编号->id 映射（设备表按插入顺序 1..N）。"""
    for i, row in enumerate(SEED_DEVICES, 1):
        if row[0] == code:
            return i
    raise ValueError(code)


# ── 示例数据：与交付 Excel《园务设备管理系统.xlsx》完全一致 ──
SEED_DEVICES = [
    ("SB-001", "教学一体机", "教学设备", "阳光园", "2022-03-15", 12800, "在用", "张莉"),
    ("SB-002", "多媒体投影仪", "教学设备", "阳光园", "2021-09-10", 4600, "在用", "张莉"),
    ("SB-003", "台式电脑", "办公设备", "阳光园", "2023-06-20", 4200, "在用", "王芳"),
    ("SB-004", "台式电脑", "办公设备", "阳光园", "2023-06-20", 4200, "维修中", "王芳"),
    ("SB-005", "激光打印机", "办公设备", "阳光园", "2022-11-05", 2300, "在用", "王芳"),
    ("SB-006", "立柜式空调", "空调暖通", "阳光园", "2020-05-18", 6800, "维修中", "李娜"),
    ("SB-007", "壁挂式空调", "空调暖通", "阳光园", "2021-04-22", 3500, "在用", "李娜"),
    ("SB-008", "餐具消毒柜", "厨房电器", "阳光园", "2022-08-30", 3200, "在用", "赵敏"),
    ("SB-009", "厨房冷藏冰箱", "厨房电器", "阳光园", "2023-02-14", 4500, "在用", "赵敏"),
    ("SB-010", "户外组合滑梯", "户外设施", "阳光园", "2019-07-08", 26800, "在用", "陈静"),
    ("SB-011", "户外攀爬架", "户外设施", "阳光园", "2019-07-08", 9600, "闲置", "陈静"),
    ("SB-012", "幼儿晨检机", "教学设备", "阳光园", "2024-03-12", 5600, "在用", "周燕"),
    ("SB-013", "教学一体机", "教学设备", "彩虹园", "2022-03-15", 12800, "在用", "张莉"),
    ("SB-014", "多媒体投影仪", "教学设备", "彩虹园", "2020-10-25", 4600, "待报废", "张莉"),
    ("SB-015", "台式电脑", "办公设备", "彩虹园", "2023-06-20", 4200, "在用", "王芳"),
    ("SB-016", "激光打印机", "办公设备", "彩虹园", "2021-12-01", 2300, "维修中", "王芳"),
    ("SB-017", "复印一体机", "办公设备", "彩虹园", "2022-04-16", 8900, "在用", "王芳"),
    ("SB-018", "立柜式空调", "空调暖通", "彩虹园", "2020-05-18", 6800, "在用", "李娜"),
    ("SB-019", "壁挂式空调", "空调暖通", "彩虹园", "2022-06-09", 3500, "闲置", "李娜"),
    ("SB-020", "餐具消毒柜", "厨房电器", "彩虹园", "2021-08-27", 3200, "在用", "赵敏"),
    ("SB-021", "和面机", "厨房电器", "彩虹园", "2023-09-06", 2800, "在用", "赵敏"),
    ("SB-022", "户外组合滑梯", "户外设施", "彩虹园", "2020-04-11", 26800, "在用", "陈静"),
    ("SB-023", "幼儿晨检机", "教学设备", "彩虹园", "2024-03-12", 5600, "在用", "周燕"),
    ("SB-024", "触控一体机", "教学设备", "星辰园", "2024-09-01", 15600, "在用", "张莉"),
    ("SB-025", "多媒体投影仪", "教学设备", "星辰园", "2021-05-19", 4600, "在用", "张莉"),
    ("SB-026", "台式电脑", "办公设备", "星辰园", "2023-06-20", 4200, "在用", "王芳"),
    ("SB-027", "针式打印机", "办公设备", "星辰园", "2019-11-13", 1800, "待报废", "王芳"),
    ("SB-028", "壁挂式空调", "空调暖通", "星辰园", "2021-04-22", 3500, "维修中", "李娜"),
    ("SB-029", "餐具消毒柜", "厨房电器", "星辰园", "2022-08-30", 3200, "在用", "赵敏"),
    ("SB-030", "紫外线消毒灯", "厨房电器", "星辰园", "2023-03-23", 900, "在用", "赵敏"),
    ("SB-031", "户外摇马组合", "户外设施", "星辰园", "2021-10-09", 7600, "闲置", "陈静"),
    ("SB-032", "教学一体机", "教学设备", "星辰园", "2024-09-01", 12800, "在用", "周燕"),
]

SEED_REPAIRS = [
    ("WX-2606-01", "SB-002", "电路故障", "2026-06-04", "2026-06-06", "已完成", 260),
    ("WX-2606-02", "SB-006", "制冷异常", "2026-06-11", "2026-06-15", "已完成", 480),
    ("WX-2606-03", "SB-020", "耗材老化", "2026-06-19", "2026-06-22", "已完成", 150),
    ("WX-2607-01", "SB-010", "机械损坏", "2026-07-03", "2026-07-10", "已完成", 1350),
    ("WX-2607-02", "SB-003", "软件故障", "2026-07-09", "2026-07-11", "已完成", 100),
    ("WX-2607-03", "SB-022", "机械损坏", "2026-07-16", "2026-07-21", "已完成", 820),
    ("WX-2607-04", "SB-009", "管路漏水", "2026-07-25", "2026-07-28", "已完成", 350),
    ("WX-2608-01", "SB-014", "光学部件", "2026-08-05", "2026-08-12", "已完成", 600),
    ("WX-2608-02", "SB-018", "制冷异常", "2026-08-13", "2026-08-16", "已完成", 420),
    ("WX-2608-03", "SB-005", "耗材老化", "2026-08-22", "2026-08-24", "已完成", 180),
    ("WX-2608-04", "SB-024", "软件故障", "2026-08-28", "2026-08-30", "已完成", 120),
    ("WX-2609-01", "SB-016", "进纸卡纸", "2026-09-02", "2026-09-05", "已完成", 220),
    ("WX-2609-02", "SB-028", "制冷异常", "2026-09-08", None, "维修中", 300),
    ("WX-2609-03", "SB-029", "电路故障", "2026-09-15", "2026-09-17", "已完成", 240),
    ("WX-2609-04", "SB-031", "机械损坏", "2026-09-21", "2026-09-26", "已完成", 310),
    ("WX-2609-05", "SB-013", "触控失灵", "2026-09-28", None, "维修中", 260),
    ("WX-2610-01", "SB-004", "软件故障", "2026-10-09", None, "待派单", 0),
    ("WX-2610-02", "SB-006", "噪音异常", "2026-10-14", "2026-10-17", "已完成", 380),
    ("WX-2610-03", "SB-027", "电路故障", "2026-10-20", None, "待派单", 0),
    ("WX-2610-04", "SB-032", "触控失灵", "2026-10-26", None, "维修中", 200),
]


def next_order_no(db, report_date_str):
    """生成下一个报修单号：WX-YYMM-XX，当月顺序递增。"""
    yymm = report_date_str[2:4] + report_date_str[5:7]
    prefix = f"WX-{yymm}-"
    rows = db.execute(
        "SELECT order_no FROM repairs WHERE order_no LIKE ?", (prefix + "%",)).fetchall()
    seq = max((int(r["order_no"][-2:]) for r in rows), default=0) + 1
    return f"{prefix}{seq:02d}"
