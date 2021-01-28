# coding:utf-8

# @Time     : 2020/10/30 9:22
# @Author   : wangzeguo
# @File     : app_old.py

from flask import Flask,render_template,request,redirect,url_for,send_from_directory,flash,session
from werkzeug.utils import secure_filename
from datetime import date
import processfile
import time
import os

from flask_bootstrap import Bootstrap
# from flask_moment import Moment
from flask_wtf import FlaskForm, Form
from wtforms.fields.html5 import DateField
from wtforms import StringField, SubmitField
from wtforms.validators import DataRequired
# from wtforms.fields import DateField
# from flask_datepicker import datepicker

app = Flask(__name__)
app.secret_key="qindklkdjfwie"
app.config['DOWNLOAD_FOLDER']='static/download/'

bootstrap = Bootstrap(app)

if not os.path.exists(app.config['DOWNLOAD_FOLDER']):
    os.makedirs(app.config['DOWNLOAD_FOLDER'])

class DateForm(FlaskForm):
    # name = StringField('What is your name?',validators=[Required()])
    date = DateField('健康码日期',default=date.today())
    submit = SubmitField('Submit')

class BookForm(Form):
    date = DateField('健康码日期', default=date.today(), validators=[DataRequired()], format='%Y/%m/%d', widget=DatePickerWidget())
    submit = SubmitField("查询")

@app.route('/info',methods=['POST','GET'])
def get_info():
    form = DateForm()
    if form.validate_on_submit():
        return render_template('/result.html',result = processfile.get_word(form.date))
    return render_template('get_info.html', form=form)


@app.route('/', methods=['POST', 'GET'])
def child():
    if request.method == 'POST':
        f = request.files['file']
        if f.filename == "":
            app.logger.info('没有选择文件')
            flash('没有选择文件')
            return redirect(url_for('child'))
        uploadpath = "static/uploads/"
        if not os.path.exists(uploadpath):
            os.makedirs(uploadpath)
        filename = time.strftime("%Y%m%d%H%M%S", time.localtime())+"_"+secure_filename(f.filename)
        f.save(uploadpath+filename)
        return render_template('result.html',result=processfile.zip_get_result(uploadpath,filename))
        # return redirect(url_for('uploaded_file', filename=processfile.get_word(uploadpath,filename)))
    return render_template('upload.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)


# @app.route('/download/<filename>')
# def download(filename):
#     return None


if __name__ == '__main__':
    app.run()