from zipfile import ZipFile
from docx import Document
from docx.shared import Inches,Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
from PIL import Image
import time

# 1.从上传的压缩包中获取图片列表并获取人名及对应的文件名
# 2.打开word，根据列表，填写题目，人名字和图片

def get_word(filepath,filename):
    doc = Document()
    doc.styles['Normal'].font.name = u'文星标宋'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'文星标宋')
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # 居中
    run = p.add_run('小一班学生及同住人健康码')
    run.font.size = Pt(18)
    img = 'static/uploads/0.jpg'  # 保存在本地的图片
    unzipfile=ZipFile(filepath+filename, 'r')
    student_list={}
    for i in reversed(unzipfile.namelist()):
        imgname =i.encode('cp437').decode('gbk')
        student_name = imgname.split('_')[1]   # 对文件名按照_切分,取出第2段
        if student_name in student_list:
            print(imgname)
            continue
        student_list[student_name]=imgname   # 给student_list元组添加学生信息
        run = doc.add_paragraph().add_run(student_name)  # 添加文字
        run.font.size = Pt(14)
        images = unzipfile.read(i)
        jpg_ima = Image.open(BytesIO(images))  # 打开图片
        jpg_ima = jpg_ima.convert('RGB')   # 去掉图片的A通道
        jpg_ima.save(img,"JPEG")  # 保存新的图片
        run = doc.add_paragraph().add_run()
        run.add_picture(img, width=Inches(3))
    print(student_list)
    downloadpath = 'static/download/'
    downloadfilename = '三明路小一班健康码'+time.strftime("%Y%m%d", time.localtime())+'.docx'
    doc.save(downloadpath+downloadfilename)  # 保存路径
    return downloadfilename