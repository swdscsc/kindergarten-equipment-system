# -*- coding: utf-8 -*-
"""路由层：管理看板、设备台账、维修记录、统计 API 与 Excel 导出。"""
import sqlite3
from datetime import datetime

from flask import (Blueprint, abort, flash, g, redirect, render_template,
                   request, url_for)

from . import db as dbm
from .excel_export import build_workbook

bp = Blueprint("main", __name__)


# ── 通用校验 ──
def _valid_date(s):
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return True
    except (TypeError, ValueError):
        return False


def _q(sql, args=()):
    return dbm.get_db().execute(sql, args).fetchall()


# ── 管理看板 ──
@bp.route("/")
def dashboard():
    db = dbm.get_db()
    total = db.execute("SELECT COUNT(*) c FROM devices").fetchone()["c"]
    by_status = {r["status"]: r["c"] for r in _q(
        "SELECT status, COUNT(*) c FROM devices GROUP BY status")}
    month_repairs = db.execute(
        "SELECT COUNT(*) c FROM repairs WHERE report_date LIKE '2026-10%'").fetchone()["c"]
    done = db.execute("SELECT COUNT(*) c FROM repairs WHERE status='已完成'").fetchone()["c"]
    all_repairs = db.execute("SELECT COUNT(*) c FROM repairs").fetchone()["c"]
    cost = db.execute("SELECT COALESCE(SUM(cost),0) s FROM repairs").fetchone()["s"]
    kpi = {
        "total": total,
        "in_use_rate": round(by_status.get("在用", 0) / total * 100, 1) if total else 0,
        "repairing": by_status.get("维修中", 0),
        "month_repairs": month_repairs,
        "done_rate": round(done / all_repairs * 100, 1) if all_repairs else 0,
        "cost": cost,
    }
    return render_template("dashboard.html", kpi=kpi)


@bp.route("/api/stats")
def api_stats():
    """看板图表数据（JSON）。"""
    parks = dbm.PARKS
    park_total = [_q("SELECT COUNT(*) c FROM devices WHERE park=?", (p,))[0]["c"] for p in parks]
    park_fixing = [_q("SELECT COUNT(*) c FROM devices WHERE park=? AND status='维修中'", (p,))[0]["c"]
                   for p in parks]
    faults = [(r["fault_type"], r["c"]) for r in _q(
        "SELECT fault_type, COUNT(*) c FROM repairs GROUP BY fault_type ORDER BY c DESC")]
    months = ["2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
    reported, finished = [], []
    for m in months:
        reported.append(_q("SELECT COUNT(*) c FROM repairs WHERE report_date LIKE ?", (m + "%",))[0]["c"])
        finished.append(_q("SELECT COUNT(*) c FROM repairs WHERE report_date LIKE ? AND status='已完成'",
                           (m + "%",))[0]["c"])
    return {"parks": parks, "parkTotal": park_total, "parkFixing": park_fixing,
            "faults": [{"name": n, "value": v} for n, v in faults],
            "months": [m[5:] + "月" for m in months], "reported": reported, "finished": finished}


# ── 设备台账 ──
@bp.route("/devices")
def devices():
    park, status, kw = request.args.get("park", ""), request.args.get("status", ""), \
        request.args.get("q", "").strip()
    sql = "SELECT * FROM devices WHERE 1=1"
    args = []
    if park:
        sql += " AND park=?"; args.append(park)
    if status:
        sql += " AND status=?"; args.append(status)
    if kw:
        sql += " AND (code LIKE ? OR name LIKE ? OR keeper LIKE ?)"
        args += [f"%{kw}%"] * 3
    sql += " ORDER BY code"
    return render_template("devices.html", rows=_q(sql, args),
                           parks=dbm.PARKS, statuses=dbm.DEVICE_STATUS,
                           f_park=park, f_status=status, f_q=kw)


def _device_form_errors(form):
    errs = []
    for f in ("code", "name", "park", "purchase_date", "value", "status", "keeper"):
        if not str(form.get(f, "")).strip():
            errs.append("必填项不能为空：" + f)
    if errs:
        return errs
    if not _valid_date(form["purchase_date"]):
        errs.append("采购日期格式应为 yyyy-mm-dd")
    try:
        if float(form["value"]) < 0:
            errs.append("采购原值不能为负数")
    except ValueError:
        errs.append("采购原值必须是数字")
    if form["status"] not in dbm.DEVICE_STATUS:
        errs.append("使用状态不合法")
    if form["park"] not in dbm.PARKS:
        errs.append("园区不合法")
    return errs


@bp.route("/devices/new", methods=["GET", "POST"])
def device_new():
    if request.method == "POST":
        errs = _device_form_errors(request.form)
        if not errs:
            exists = _q("SELECT 1 FROM devices WHERE code=?", (request.form["code"].strip(),))
            if exists:
                errs.append("设备编号已存在：" + request.form["code"].strip())
        if errs:
            for e in errs:
                flash(e, "error")
        else:
            f = request.form
            dbm.get_db().execute(
                "INSERT INTO devices (code,name,category,park,purchase_date,value,status,keeper)"
                " VALUES (?,?,?,?,?,?,?,?)",
                (f["code"].strip(), f["name"].strip(), f["category"], f["park"],
                 f["purchase_date"], float(f["value"]), f["status"], f["keeper"].strip()))
            dbm.get_db().commit()
            flash("设备已新增：" + f["code"].strip(), "ok")
            return redirect(url_for("main.devices"))
    return render_template("device_form.html", d=request.form, parks=dbm.PARKS,
                           statuses=dbm.DEVICE_STATUS, categories=dbm.CATEGORIES,
                           title="新增设备")


@bp.route("/devices/<int:did>/edit", methods=["GET", "POST"])
def device_edit(did):
    row = _q("SELECT * FROM devices WHERE id=?", (did,))
    if not row:
        abort(404)
    if request.method == "POST":
        errs = _device_form_errors(request.form)
        if not errs:
            dup = _q("SELECT 1 FROM devices WHERE code=? AND id<>?",
                     (request.form["code"].strip(), did))
            if dup:
                errs.append("设备编号已存在：" + request.form["code"].strip())
        if errs:
            for e in errs:
                flash(e, "error")
        else:
            f = request.form
            dbm.get_db().execute(
                "UPDATE devices SET code=?,name=?,category=?,park=?,purchase_date=?,"
                "value=?,status=?,keeper=? WHERE id=?",
                (f["code"].strip(), f["name"].strip(), f["category"], f["park"],
                 f["purchase_date"], float(f["value"]), f["status"], f["keeper"].strip(), did))
            dbm.get_db().commit()
            flash("设备已更新：" + f["code"].strip(), "ok")
            return redirect(url_for("main.devices"))
    return render_template("device_form.html", d=row[0] if request.method == "GET" else request.form,
                           parks=dbm.PARKS, statuses=dbm.DEVICE_STATUS,
                           categories=dbm.CATEGORIES, title="编辑设备")


@bp.route("/devices/<int:did>/delete", methods=["POST"])
def device_delete(did):
    used = _q("SELECT COUNT(*) c FROM repairs WHERE device_id=?", (did,))[0]["c"]
    if used:
        flash(f"该设备存在 {used} 条维修记录，禁止删除；如需处置请将使用状态改为「待报废」。", "error")
    else:
        dbm.get_db().execute("DELETE FROM devices WHERE id=?", (did,))
        dbm.get_db().commit()
        flash("设备已删除。", "ok")
    return redirect(url_for("main.devices"))


# ── 维修记录 ──
@bp.route("/repairs")
def repairs():
    status, park = request.args.get("status", ""), request.args.get("park", "")
    sql = ("SELECT r.*, d.code dcode, d.name dname, d.park park FROM repairs r "
           "JOIN devices d ON d.id=r.device_id WHERE 1=1")
    args = []
    if status:
        sql += " AND r.status=?"; args.append(status)
    if park:
        sql += " AND d.park=?"; args.append(park)
    sql += " ORDER BY r.report_date DESC, r.order_no DESC"
    return render_template("repairs.html", rows=_q(sql, args),
                           statuses=dbm.REPAIR_STATUS, parks=dbm.PARKS,
                           f_status=status, f_park=park)


@bp.route("/repairs/new", methods=["GET", "POST"])
def repair_new():
    if request.method == "POST":
        f = request.form
        errs = []
        dev = _q("SELECT * FROM devices WHERE code=?", (f.get("device_code", "").strip(),))
        if not dev:
            errs.append("设备编号不存在，请核对台账（例：SB-001）")
        if not _valid_date(f.get("report_date")):
            errs.append("报修日期格式应为 yyyy-mm-dd")
        if f.get("fault_type") not in dbm.FAULT_TYPES:
            errs.append("请从下拉列表选择标准故障类型")
        if not errs:
            db = dbm.get_db()
            order_no = dbm.next_order_no(db, f["report_date"])
            db.execute(
                "INSERT INTO repairs (order_no,device_id,fault_type,report_date,finish_date,status,cost)"
                " VALUES (?,?,?,?,NULL,'待派单',0)",
                (order_no, dev[0]["id"], f["fault_type"], f["report_date"]))
            if dev[0]["status"] == "在用":
                db.execute("UPDATE devices SET status='维修中' WHERE id=?", (dev[0]["id"],))
            db.commit()
            flash(f"报修单已登记：{order_no}（设备 {dev[0]['code']}）", "ok")
            return redirect(url_for("main.repairs"))
        for e in errs:
            flash(e, "error")
    return render_template("repair_form.html", faults=dbm.FAULT_TYPES,
                           today=datetime.now().strftime("%Y-%m-%d"), f=request.form)


@bp.route("/repairs/<int:rid>/update", methods=["POST"])
def repair_update(rid):
    rows_ = _q("SELECT r.*, d.id device_pk, d.status dstatus FROM repairs r "
               "JOIN devices d ON d.id=r.device_id WHERE r.id=?", (rid,))
    if not rows_:
        abort(404)
    row = rows_[0]
    f = request.form
    errs = []
    status = f.get("status")
    if status not in dbm.REPAIR_STATUS:
        errs.append("维修状态不合法")
    if not errs and status == "已完成":
        finish = f.get("finish_date", "").strip()
        if not _valid_date(finish):
            errs.append("标记为已完成时必须填写有效的完成日期（yyyy-mm-dd）")
        try:
            cost = float(f.get("cost", 0))
            if cost < 0:
                errs.append("维修费用不能为负数")
        except ValueError:
            errs.append("维修费用必须是数字")
    if errs:
        for e in errs:
            flash(e, "error")
        return redirect(url_for("main.repairs"))

    db = dbm.get_db()
    finish = f.get("finish_date", "").strip() or None
    cost = float(f.get("cost", 0) or 0)
    if status != "已完成":
        finish, cost = row["finish_date"], row["cost"]
    db.execute("UPDATE repairs SET status=?, finish_date=?, cost=? WHERE id=?",
               (status, finish, cost, rid))
    # 状态联动：维修单未完成 → 设备为维修中；维修完成 → 设备恢复在用
    if status in ("待派单", "维修中"):
        db.execute("UPDATE devices SET status='维修中' WHERE id=?", (row["device_pk"],))
    elif status == "已完成" and row["dstatus"] == "维修中":
        db.execute("UPDATE devices SET status='在用' WHERE id=?", (row["device_pk"],))
    db.commit()
    flash(f"维修单 {row['order_no']} 已更新为「{status}」。", "ok")
    return redirect(url_for("main.repairs"))


@bp.route("/repairs/<int:rid>/delete", methods=["POST"])
def repair_delete(rid):
    db = dbm.get_db()
    db.execute("DELETE FROM repairs WHERE id=?", (rid,))
    db.commit()
    flash("维修记录已删除。", "ok")
    return redirect(url_for("main.repairs"))


# ── Excel 导出 ──
@bp.route("/export.xlsx")
def export():
    wb = build_workbook(dbm.get_db())
    import io
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    stamp = datetime.now().strftime("%Y%m%d")
    return (buf.getvalue(), 200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": f"attachment; filename=equipment_{stamp}.xlsx",
    })
