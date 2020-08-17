# coding:utf-8

from flask import Flask,render_template,request,redirect,url_for,send_from_directory
from werkzeug.utils import secure_filename
import processfile
import time
import os

app = Flask(__name__)
app.config['DOWNLOAD_FOLDER']='static/download/'

if not os.path.exists(app.config['DOWNLOAD_FOLDER']):
    os.makedirs(app.config['DOWNLOAD_FOLDER'])

@app.route('/', methods=['POST', 'GET'])
def upload():
    if request.method == 'POST':
        f = request.files['file']
        if f.filename == "":
            app.logger.info('没有选择文件')
            return redirect(url_for(''))
        uploadpath = "static/uploads/"
        if not os.path.exists(uploadpath):
            os.makedirs(uploadpath)
        filename = time.strftime("%Y%m%d%H%M%S", time.localtime())+"_"+secure_filename(f.filename)
        f.save(uploadpath+filename)
        return redirect(url_for('uploaded_file', filename=processfile.get_word(uploadpath,filename)))
    return render_template('upload.html')

@app.route('/download/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run()