import os
import re
import sys
import win32com.client
import pythoncom
import flask, json
from flask import request, render_template

from finereport_docx_process import doc_process

server = flask.Flask(__name__)
root_path = os.getcwd()
# 保存上传上来的文件
UPLOAD_FOLDER = root_path + '\\UPLOADS'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def allowed_file(filename):
    ALLOWED_EXTENSIONS = ['docx']
    # 判断文件的扩展名是否在配置项ALLOWED_EXTENSIONS中
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

@server.route('/upload', methods=['get', 'post'])
def upload():
        用于对已经在服务器内的文件进行操作，提供的是文件的地址
    # 获取通过url请求传参的数据
    file = request.values.get('source')
    files = file.split(',')
    # 获取url请求传的密码，明文
    target = request.values.get('tar')
    # 判断用户名、密码都不为空，如果不传用户名、密码则username和pwd为None
    if len(files) != 2:
        resu = {'code': 400, 'message': '需要传递文件和表格的文件位置，用半角的,分隔！'}
        return json.dumps(resu, ensure_ascii=False)
    elif type(target) != str:
        resu = {'code': 400, 'message': '需要存放生成文件的位置'}
        return json.dumps(resu, ensure_ascii=False)
    else:
        target_path, tar_file = os.path.split(target)
        if not os.path.exists(target_path):
            os.makedirs(target_path)

        if '表格' in files[0]:
            doc_process(files[1], files[0], target).run()
        else:
            doc_process(files[0], files[1], target).run()
        resu = {'code': 200, 'message': '转换成功'}
        return json.dumps(resu, ensure_ascii=False)
        # if request.method == 'POST':
    #     # 获取上传过来的文件对象
    #     file = request.files['file']
    #     # 检查文件对象是否存在，且文件名合法
    #     if file and allowed_file(file.filename):
    #         # 去除文件名中不合法的内容
    #         # 将文件保存在本地UPLOAD_FOLDER目录下
    #         file.save(os.path.join(UPLOAD_FOLDER, file.filename))

    #         resu = {'code': 200, 'message': '上传成功'}
    #         return json.dumps(resu, ensure_ascii=False)
    #     else:    # 文件不合法
    #         resu = {'code': 400, 'message': '文件格式有误'}
    #         return json.dumps(resu, ensure_ascii=False)
    # else:  # GET方法
    #     return render_template('upload.html')

if __name__ == '__main__':
    server.run(port=7777, debug=True, host='0.0.0.0')
# http://127.0.0.1:7777/process_doc?source=D:\fineReports\software\2.0%E8%BD%AF%E4%BB%B6\2020\6\%E6%B5%B7%E5%8D%97%E7%9C%81\word\2.0%E9%80%9A%E4%BF%A1%E5%9B%BE%E6%96%87.docx,C:\Users\iceberg\Desktop\%E8%87%AA%E5%8A%A8%E5%8C%96word\%E8%A1%A8%E6%A0%BC.docx&tar=C:\Users\iceberg\Desktop\%E8%87%AA%E5%8A%A8%E5%8C%96word\%E5%A4%84%E7%90%86%E7%BB%93%E6%9E%9C\res.docx
