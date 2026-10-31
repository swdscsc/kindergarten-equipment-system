# -*- coding: utf-8 -*-
"""启动入口：python run.py [init-db]"""
import sys

from app import create_app

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "init-db":
        application = create_app()
        print("数据库已初始化：", application.db_path)
    else:
        port = int(sys.argv[2]) if len(sys.argv) > 2 else 5177
        application = create_app()
        print(f"园务设备管理系统已启动：http://127.0.0.1:{port}")
        application.run(debug=False, port=port)
