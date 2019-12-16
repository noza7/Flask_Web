import os

DEBUG = True
SECRET_KEY = os.urandom(24)
SQLALCHEMY_DATABASE_URI = 'sqlite:///db.sqlite3'

# 关闭数据跟踪修改
SQLALCHEMY_TRACE_MODIFICATIONS = False
