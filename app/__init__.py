# -*- coding: utf-8 -*-
"""园务设备管理系统 — Flask 应用工厂。"""
import os
import sqlite3

from flask import Flask, g, redirect, render_template, request

from . import db as dbm
from .routes import bp


def create_app(instance_path=None):
    app = Flask(__name__, instance_relative_config=True,
                instance_path=instance_path)
    app.config["SECRET_KEY"] = "kindergarten-equipment-demo"

    # 首次启动自动建表并写入示例数据（示例数据与交付 Excel 一致）
    dbm.init_db(app, seed=True)
    app.config["DB_PATH"] = dbm.db_path(app)

    app.register_blueprint(bp)

    @app.teardown_appcontext
    def _close(e=None):
        dbm.close_db(e)

    @app.errorhandler(404)
    def not_found(e):
        return render_template("error.html", msg="页面不存在（404）"), 404

    @app.errorhandler(sqlite3.Error)
    def db_err(e):
        if request.method == "POST":
            return redirect(request.referrer or "/", code=302)
        return render_template("error.html", msg="数据操作失败，请检查输入"), 500

    return app
