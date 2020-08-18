from zipfile import ZipFile
from docx import Document
from docx.shared import Inches,Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
from PIL import Image
import time
from config import config

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


def zip_get_result(filepath,filename):
    # 生成word文档标题
    doc = Document()
    doc.styles['Normal'].font.name = u'文星标宋'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'文星标宋')
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # 居中
    run = p.add_run('小一班学生及同住人健康码')
    run.font.size = Pt(18)
    img = 'static/uploads/0.jpg'  # 保存在本地的图片
    # 获取提交学生名单至student_list
    unzipfile=ZipFile(filepath+filename, 'r')
    all_student = config.get('classone').get('student_name') # 全部学生名单
    student_list = {} # 已提交名单
    repeat_list = {} # 重复提交名单
    undo_list = {} # 未提交名单
    for i in reversed(unzipfile.namelist()):
        imgname =i.encode('cp437').decode('gbk')
        student_filename = imgname.split('_')  # 对文件名按照_切分,取出第2段
        student_name = student_filename[1]
        all_student[student_name]=1
        if student_name in student_list:
            repeat_list[student_filename[0]]=student_name
            # print(imgname)
            continue
        student_list[student_name]=student_filename[0]   # 给student_list元组添加学生信息
        run = doc.add_paragraph().add_run(student_name)  # 添加文字
        run.font.size = Pt(14)
        images = unzipfile.read(i)
        jpg_ima = Image.open(BytesIO(images))  # 打开图片
        jpg_ima = jpg_ima.convert('RGB')   # 去掉图片的A通道
        jpg_ima.save(img,"JPEG")  # 保存新的图片
        run = doc.add_paragraph().add_run()
        run.add_picture(img, width=Inches(3))
    # 保存word文档
    downloadpath = 'static/download/'
    downloadfilename = '三明路小一班健康码'+time.strftime("%Y%m%d", time.localtime())+'.docx'
    doc.save(downloadpath+downloadfilename)  # 保存路径
    # 计算未提交人员名单
    for i in all_student:
        print(i,all_student[i])
        if all_student[i] == '0':
            undo_list[i] = all_student[i]
    result_list = {} #返回结果至页面
    result_list['student'] = student_list
    result_list['repeat'] = repeat_list
    result_list['undo'] = undo_list
    result_list['filename'] = downloadfilename
    result_list['student_size'] = len(student_list)
    result_list['repeat_size'] = len(repeat_list)
    result_list['undo_size'] = len(undo_list)
    return result_list