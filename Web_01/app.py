from datetime import timedelta

from flask import Flask, render_template, request, redirect, url_for, session, flash

import config
from exts import db
from models import User

from functools import wraps

app = Flask(__name__)
# 设置静态文件缓存过期时间
# app.config['SEND_FILE_MAX_AGE_DEFAULT'] = timedelta(seconds=1)
app.send_file_max_age_default = timedelta(seconds=1)
app.config.from_object(config)
db.init_app(app)


# 登录限制装饰器
def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if session.get('user_id'):
            return func(*args, **kwargs)
        else:
            return redirect(url_for('login'))

    return wrapper


@app.route('/')
@login_required
def index():
    return render_template('index.html')


@app.route('/login/', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template('login.html')
    else:
        tel = request.form.get('tel')
        password = request.form.get('password')
        # 同时查询用户手机号和密码
        user = User.query.filter(User.tel == tel, User.password == password).first()
        # 如果存在
        if user:
            # 保存session到cookie
            session['user_id'] = user.id
            # 31天不用重新登录
            session.permanent = True
            return redirect(url_for('index'))
        else:
            flash('手机号码或密码不正确！', category='error')
            return render_template('login.html')


@app.route('/logout/')
def logout():
    # session.pop('user_id')
    session.clear()
    return redirect(url_for('login'))


@app.route('/register/', methods=['GET', 'POST'])
def register():
    if request.method == 'GET':
        return render_template('register.html')
    else:
        tel = request.form.get('tel')
        username = request.form.get('username')
        password = request.form.get('password')
        password_confirm = request.form.get('password_confirm')

        # 验证手机号码是否被注册
        user_tel = User.query.filter(User.tel == tel).first()
        # 判断手机号码长度是否正确
        if len(tel) != 11:
            flash('手机号码长度不正确！', category='error')
            return render_template('register.html')
        if user_tel:
            flash('该手机号码已经注册！', category='error')
            return render_template('register.html')
            # return '该手机号码已经注册！'
        else:
            # 判断密码长度
            if len(password) < 8:
                flash('密码长度过短！', category='error')
                return render_template('register.html')
            # 判断两次输入密码是否相等
            if password != password_confirm:
                flash('两次密码输入不同', category='error')
                return render_template('register.html')
                # return '两次密码输入不同！'
            else:
                # 如果都没问题，创建新用户到数据库
                user = User(tel=tel, username=username, password=password)
                db.session.add(user)
                db.session.commit()
                # 注册成功，跳转到登录页面
                return redirect(url_for('login'))


# 上下文处理器
@app.context_processor
def my_context_processor():
    user_id = session.get('user_id')
    if user_id:
        user = User.query.filter(User.id == user_id).first()
        if user:
            return {'user': user}
    return {}


if __name__ == '__main__':
    app.run()
