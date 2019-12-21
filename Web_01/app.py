import os
from datetime import timedelta

from flask import Flask, render_template, request, redirect, url_for, session, flash, make_response, send_from_directory

import config
from exts import db, creat_folder, delfile, multi_tables
from models import User

from functools import wraps

from flask_uploads import UploadSet, DOCUMENTS, configure_uploads

app = Flask(__name__)
# 设置静态文件缓存过期时间
# app.config['SEND_FILE_MAX_AGE_DEFAULT'] = timedelta(seconds=1)
app.send_file_max_age_default = timedelta(seconds=1)
app.config.from_object(config)
db.init_app(app)

# 文件上传配置项
# 上传文件的最大尺寸
# app.config['MAX_CONTENT_LENGTH'] = 16 * 24 * 24
# 上传路径配置
APPS_DIR = os.path.dirname(__file__)  # 获取app路径名
STATIC_DIR = os.path.join(APPS_DIR, 'static')  # 通过拼接获取 'static'路径名
app.config['UPLOAD_FOLDER'] = 'uploads'  # 配置上传文件夹为'uploads'
app.config['ABS_UPLOAD_FOLDER'] = os.path.join(STATIC_DIR, app.config['UPLOAD_FOLDER'])  # 设置上传绝对路径
creat_folder(app.config['ABS_UPLOAD_FOLDER'])


def upload_set(path, set_name, FILE_TYPE):
    '''
    # 定义 Upload Sets 配置
    :param path: 要保存文件的路径
    :param set_name: 要创建的集合名称
    :param FILE_TYPE: 要保存的文件类型
    :return:返回app.config路径 和 要管理上传的集合实例
    '''
    # 设置文件类型保存地址
    app.config[f'UPLOADED_{set_name.upper()}_DEST'] = app.config['ABS_UPLOAD_FOLDER'] + path
    app_path = app.config[f'UPLOADED_{set_name.upper()}_DEST']
    set_name = UploadSet(f'{set_name.lower()}', FILE_TYPE)
    return app_path, set_name


# Upload Sets,管理上传集合
chengzhao_app_path, chengzhao_set = upload_set(path=r'\chengzhao\temp', set_name='chengzhao', FILE_TYPE=DOCUMENTS)
# 绑定 app 与 UploadSets
configure_uploads(app, chengzhao_set)


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


# 成招成绩单文件上传
@app.route('/chengzhao/', methods=['GET', 'POST'])
@login_required
def chengzhao():
    if request.method == 'POST':
        f = request.files['excel_upload']
        if f.filename != '':
            try:
                f_name = chengzhao_set.save(f)
                file_url = chengzhao_set.url(f_name)
                flash(f'文件{f_name}上传成功！')
                return render_template('chengzhao.html', file_url=file_url)
            except Exception as e:
                print(e)
                flash('请检查文件格式！', category='error')
                return render_template('chengzhao.html')
    elif request.method == 'GET':
        print(chengzhao_app_path)
        path = chengzhao_app_path  # 文件路径
        if os.path.exists(path):  # 如果文件存在
            delfile(path)
        else:
            print('没有文件！')  # 则返回文件不存在
    return render_template('chengzhao.html')


# 成招成绩单生成文件下载
@app.route('/chengzhaodownload/', methods=['GET'])
@login_required
def chengzhaodownload():
    try:
        multi_tables(chengzhao_app_path)
        filename = 'result.xlsx'
        directory = chengzhao_app_path
        response = make_response(send_from_directory(directory, filename, as_attachment=True))
        return response
    except Exception as e:
        print(e)
        flash('文件上传有误，请检查后重新上传！', category='error')
        return render_template('chengzhao.html')


# 登录
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


# 注销
@app.route('/logout/')
def logout():
    # session.pop('user_id')
    session.clear()
    return redirect(url_for('login'))


# 注册
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
