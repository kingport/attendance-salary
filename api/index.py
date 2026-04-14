"""
Vercel Serverless 入口
薄壳：直接复用项目根目录的 app.py，避免两套代码漂移
"""
import os
import sys

# 将项目根目录加入 sys.path，以便 import app
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app  # noqa: F401,E402  Vercel 会找名为 app 的 WSGI 对象
