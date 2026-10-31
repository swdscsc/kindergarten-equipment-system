# -*- coding: utf-8 -*-
"""pytest 公共夹具：每个用例使用独立的临时数据库。"""
import tempfile

import pytest

from app import create_app


@pytest.fixture()
def app(tmp_path):
    application = create_app(instance_path=str(tmp_path / "instance"))
    application.config["TESTING"] = True
    yield application


@pytest.fixture()
def client(app):
    return app.test_client()
