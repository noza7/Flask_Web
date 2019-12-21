import os

DEBUG = True
SECRET_KEY = os.urandom(24)
SQLALCHEMY_DATABASE_URI = 'sqlite:///db.sqlite3'

# 关闭数据跟踪修改
SQLALCHEMY_TRACE_MODIFICATIONS = False

# 文件上传
# 上传文件的最大尺寸
# MAX_CONTENT_LENGTH = 16 * 24 * 24  # 这里最大是16M

