# coding:utf-8

# @Time     : 2021/1/25 17:27
# @Author   : wangzeguo
# @File     : app.py.py

from flask import Flask
from flask import render_template,send_from_directory
from flask_bootstrap import Bootstrap
from flask import request
import getinfofromwjx

app = Flask(__name__,static_url_path='')
app.config['BOOTSTRAP_SERVE_LOCAL']=True
app.config['SECRET_KEY'] = 'hard to guess string'
app.config['DOWNLOAD_FOLDER']='static/download/'
bootstrap = Bootstrap(app)

@app.route('/')
def index():
    return render_template('index.html',name = '三明路')

@app.route('/result')
def getinfo():
    resultdate = request.args.get('start')
    return render_template('get_info.html',date = resultdate,result = getinfofromwjx.get_excle(resultdate))

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run()
