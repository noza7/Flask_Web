import os
from datetime import timedelta

from flask import Flask, render_template, request, redirect, url_for, session, flash

import config
from exts import db, creat_folder, chengzhao_output_scores_tables, to_zip, part_print_do, class_arrange_do
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
# 成招临时文件集合
chengzhao_temp_app_path, chengzhao_temp_set = upload_set(path=r'\chengzhao\temp', set_name='chengzhaotemp',
                                                         FILE_TYPE=DOCUMENTS)
# 绑定 app 与 UploadSets（成招临时文件）
configure_uploads(app, chengzhao_temp_set)

# 开放教育临时文件集合
kfjy_temp_app_path, kfjy_temp_set = upload_set(path=r'\kfjy\temp', set_name='kfjytemp', FILE_TYPE=DOCUMENTS)
# 绑定 app 与 UploadSets（成招临时文件）
configure_uploads(app, kfjy_temp_set)

# 模板文件集合
templatefiles_path, templatefiles_set = upload_set(path=r'\templatefiles', set_name='templatefiles',
                                                   FILE_TYPE=DOCUMENTS)
# 绑定 app 与 UploadSets（模板）
configure_uploads(app, templatefiles_set)


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


# 成招成绩单文件上传，生成
@app.route('/chengzhao/', methods=['GET', 'POST'])
@login_required
def chengzhao():
    scores_path = templatefiles_set.url('scores.xlsx')
    template_path = templatefiles_set.url('template.xlsx')
    if request.method == 'POST':
        f = request.files['excel_upload']
        if f.filename != '':
            try:
                f_name = chengzhao_temp_set.save(f)
                file_url = chengzhao_temp_set.url(f_name)
                flash(f'文件{f_name}上传成功！')
                return render_template('chengzhao.html', file_url=file_url)
            except Exception as e:
                print(e)
                flash('请检查文件格式！', category='error')
                return render_template('chengzhao.html', scores_path=scores_path, template_path=template_path)
    elif request.method == 'GET':
        path = chengzhao_temp_app_path  # 文件路径
        print(path)
        if os.path.exists(path):  # 如果文件存在
            # 因为需要权限才能进行删除操作，所以只能用这种方法来进行目录的删除操作
            os.system(f"rd/s/q  {path}")
            # rmtree(path)
            # delfile(output_path)
            # delfile(path)
        else:
            print('没有文件！')  # 则返回文件不存在
    return render_template('chengzhao.html', scores_path=scores_path, template_path=template_path)


# 成招成绩单处理
@app.route('/chengzhaodownload/', methods=['GET'])
@login_required
def chengzhaodownload():
    try:
        # 文件处理
        # 成绩单输出路径
        output_path = chengzhao_temp_app_path + r'\output'
        flash('正在处理中，请稍后......')
        chengzhao_output_scores_tables(chengzhao_temp_app_path)
        flash('正在压缩，请稍后......')
        # 压缩文件夹
        to_zip(output_path)
        # 要返回的文件路径
        output_file_url = chengzhao_temp_set.url('output.zip')
        # 删除 output 文件夹
        # delfile(output_path)
        # print(output_file_url)
        # 文件下载
        # filename = 'output.zip'
        # directory = chengzhao_app_path
        # response = make_response(send_from_directory(directory, filename, as_attachment=True))
        return render_template('chengzhao.html', output_file_url=output_file_url)
    except Exception as e:
        print(e)
        flash('文件上传有误，请检查后重新上传！', category='error')
        return render_template('chengzhao.html')


# 开放教育
@app.route('/kfjy/', methods=['GET', 'POST'])
@login_required
def kfjy():
    return render_template('kfjy/kfjy.html')


# 打印门贴签到表
@app.route('/PartPrint/', methods=['GET', 'POST'])
@login_required
def partprint():
    kcqkb_path = templatefiles_set.url('kcqkb.xls')
    if request.method == 'POST':
        f = request.files['excel_upload']
        if f.filename != '':
            try:
                f_name = kfjy_temp_set.save(f)
                file_url = kfjy_temp_set.url(f_name)
                flash(f'文件{f_name}上传成功！')
                return render_template('kfjy/PartPrint.html', file_url=file_url, kcqkb_path=kcqkb_path)
            except Exception as e:
                print(e)
                flash('请检查文件格式！', category='error')
                return render_template('kfjy/PartPrint.html', kcqkb_path=kcqkb_path)
    elif request.method == 'GET':
        path = kfjy_temp_app_path  # 文件路径
        print(path)
        if os.path.exists(path):  # 如果文件存在
            # 因为需要权限才能进行删除操作，所以只能用这种方法来进行目录的删除操作
            os.system(f"rd/s/q  {path}")
        else:
            print('没有文件！')  # 则返回文件不存在
    return render_template('kfjy/PartPrint.html', kcqkb_path=kcqkb_path)


# 打印门贴签到表处理
@app.route('/PartPrintDo/', methods=['GET'])
@login_required
def partprintdo():
    try:
        # 文件处理
        # 成绩单输出路径
        file_path1 = kfjy_temp_app_path + r'\kcqkb.xls'
        file_path2 = kfjy_temp_app_path + r'\data.txt'
        flash('正在处理中，请稍后......')
        part_print_do(path=file_path1, my_file=file_path2)
        # 要返回的文件路径
        output_file_url = kfjy_temp_set.url('data.txt')
        flash('文件处理完成，请下载！')
        return render_template('kfjy/PartPrint.html', output_file_url=output_file_url)
    except Exception as e:
        print(e)
        flash('文件上传有误，请检查后重新上传！', category='error')
        return render_template('kfjy/PartPrint.html')


# 开放教育课程编排
@app.route('/ClassArrange/', methods=['GET', 'POST'])
@login_required
def class_arrange():
    class_info_path = templatefiles_set.url('class_info.xlsx')
    class_arrange_path = templatefiles_set.url('class_arrange.xlsx')
    if request.method == 'POST':
        f = request.files['excel_upload']
        f2 = request.files['excel_upload2']
        # 通过定义全局变量获取教室数
        global g_classroom_num
        g_classroom_num = request.form.get('classroom_num')
        print(g_classroom_num)
        if f.filename != '' and f2.filename != '':
            try:
                f_name = kfjy_temp_set.save(f)
                file_url = kfjy_temp_set.url(f_name)
                f_name2 = kfjy_temp_set.save(f2)
                file_url2 = kfjy_temp_set.url(f_name2)
                flash(f'文件{f_name}上传成功！')
                flash(f'文件{f_name2}上传成功！')
                return render_template('kfjy/ClassArrange.html', file_url=file_url, file_url2=file_url2,class_info_path=class_info_path,
                                       class_arrange_path=class_arrange_path)
            except Exception as e:
                print(e)
                flash('请检查文件格式！', category='error')
                return render_template('kfjy/ClassArrange.html', class_info_path=class_info_path,
                                       class_arrange_path=class_arrange_path)
    elif request.method == 'GET':
        path = kfjy_temp_app_path  # 文件路径
        print(path)
        if os.path.exists(path):  # 如果文件存在
            # 因为需要权限才能进行删除操作，所以只能用这种方法来进行目录的删除操作
            os.system(f"rd/s/q  {path}")
        else:
            print('没有文件！')  # 则返回文件不存在
        return render_template('kfjy/ClassArrange.html', class_info_path=class_info_path,
                               class_arrange_path=class_arrange_path)
    return render_template('kfjy/ClassArrange.html', class_info_path=class_info_path,
                           class_arrange_path=class_arrange_path)


# 开放教育课程编排处理程序
@app.route('/KfClassArrangeProcess/', methods=['GET', 'POST'])
@login_required
def class_arrange_process():
    try:
        # 教室数
        CLASSROOM_NUM = int(g_classroom_num)
        # 文件处理
        # 输出路径
        class_info_path = kfjy_temp_app_path + r'\class_info.xlsx'
        class_arrange_path = kfjy_temp_app_path + r'\class_arrange.xlsx'
        output_path = kfjy_temp_app_path + '\\'
        flash('正在处理中，请稍后......')
        # 教室数
        # CLASSROOM_NUM = 14
        class_arrange_do(path=class_arrange_path, CLASSROOM_NUM=CLASSROOM_NUM, class_info_path=class_info_path,
                         output_path=output_path)
        # 要返回的文件路径
        output_file_url = kfjy_temp_set.url('排课结果.xlsx')
        flash('文件处理完成，请下载！')
        return render_template('kfjy/ClassArrangeProcess.html', output_file_url=output_file_url)
    except Exception as e:
        print(e)
        flash('文件上传有误，请检查后重新上传！', category='error')
        return render_template('kfjy/ClassArrangeProcess.html')


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
    app.run(
        host='0.0.0.0',
        port=80,
        debug=True
    )