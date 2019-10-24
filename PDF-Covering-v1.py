# -*- coding: UTF-8 -*-
import os
import gc
import shutil
import importlib, sys
# importlib.reload(sys)

from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import *
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
import multiprocessing
from multiprocessing import freeze_support

file_list = []
file_name = []
read_dir = 'sop'
output_dir = 'output'
watermark_dir = "watermark"


def check_dir(dir_name):
    # 判断路径dir_name是否存在
    if os.path.exists(dir_name):
        print("目录:"+dir_name+"存在")
    else:
        os.mkdir("./"+dir_name)
        print("已创建目录"+"./"+dir_name)


def remove_file(path_file):
    if os.path.exists(path_file):
        os.remove(path_file)


def remove_dir(dir_remove):
    if os.path.exists(dir_remove):
        shutil.rmtree(dir_remove)


def get_pdf_namelist(path):
    # path 是读取pdf文件的文件夹,返回[pdf文件名列表，pdf文件路径列表]
    check_dir(path)
    pdfpathlist = []
    pdfnamelist = []
    for root, dirs, files in os.walk(path):
        for file in files:
            file_tumple = os.path.splitext(file)
            filename, filetype = file_tumple
            if filetype == '.pdf':
                pdf_file_path = "./"+path+"/"+file
                pdfpathlist.append(pdf_file_path)
                pdfnamelist.append(file)

    return [pdfnamelist,pdfpathlist]
# 创建透明pdf文件
# x,y pdf页面长宽


def create_watermark(x, y):
    # 创建透明pdf文件模板
    # x,y pdf页面长宽
    # 新建一个默认空白文件
    c = canvas.Canvas("mark-none.pdf", pagesize=(x, y))
    # 移动坐标原点(坐标系左下为(0,0))
    # 设置字体
    c.setFont("Helvetica", 0)
    # 指定描边的颜色
    c.setStrokeColorRGB(1, 1, 1)
    # 指定填充颜色
    c.setFillColorRGB(1, 1, 1)
    # 画一个矩形
    c.rect(0, 0, 0, 0, fill=1)
    # 关闭并保存pdf文件
    c.save()


def create_watermark_pdf(input_pdf_name, watermark):
    # input_pdf_name 原pdf文件名， output_dir  输出路径，  watermark 水印文件名
    #读取水印页面
    watermark_obj = PdfFileReader(watermark)
    watermark_page = watermark_obj.getPage(0)

    #指定原pdf文件
    in_pdf_path = "./" + read_dir + "/" + input_pdf_name
    print(read_dir)
    print(input_pdf_name)
    print(input_pdf_name)
    pdf_reader = PdfFileReader(in_pdf_path)
    pdf_writer = PdfFileWriter()

    #循环每一页，插入空白页面
    cur_page = 0
    for page in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page)
        page.mergePage(watermark_page)
        pdf_writer.addPage(page)

        # 打印进度
        cur_page += 1
        #获取完成的页数
        rate = float(cur_page)/float(pdf_reader.getNumPages())
        #输入完成的进度
        print('File conversion completed :'+'%.2f%%' % (rate * 100)+'.')

    #完成输出文件,使用全局变量watermark_dir ，将水印文件保存到某目录
    water_output_dir = watermark_dir
    check_dir(water_output_dir)
    # 保存到水印目录下，名字与原pdf文件同名
    output = "./"+water_output_dir+"/" + input_pdf_name
    with open(output, 'wb') as out:
        pdf_writer.write(out)
        # 打印创建水印结果成功
        print('The '+input_pdf_name+' file first add watermark successful.')


def copy_to_other_dir(filename,other_dir):
    path_src = "./" + read_dir + "/" + filename
    path_dst = "./" + other_dir + "/" + filename
    shutil.copy(path_src,path_dst)


# 解析pdf
def parse(pdf_file, code_source, code_new):
    pdf_path_ = "./"+watermark_dir+"/" + pdf_file
    with open(pdf_path_, 'rb') as pdf_io:
    # 用文件对象创建一个PDF文档分析器
    # parser = PDFParser(DataIO)
        parser = PDFParser(pdf_io)
    # 创建一个PDF文档
    doc = PDFDocument()
    # 分析器和文档相互连接
    parser.set_document(doc)
    doc.set_parser(parser)
    # 提供初始化密码，没有默认为空
    doc.initialize()
    # 检查文档是否可以转成文本，如果不可以读取文本，抛出异常
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        # 创建PDF资源管理器，来管理共享资源
        rsrcmagr = PDFResourceManager()
        # 创建一个PDF设备对象
        laparams = LAParams()
        # 将资源管理器和设备对象聚合
        device = PDFPageAggregator(rsrcmagr, laparams=laparams)
        # 创建一个PDF解释器对象
        interpreter = PDFPageInterpreter(rsrcmagr, device)

        pg = 0
        # 循环遍历列表，每次处理一个page内容
        # print(doc.get_pages())
        # 获取page列表
        for page in doc.get_pages():
            size_x = []
            size_y = []
            size_font = []
            layout_x = 2448
            layout_y = 1584
            try:
                interpreter.process_page(page)
                # 接收该页面的LTPage对象
                layout = device.get_result()
                # 这里的layout是一个LTPage对象 里面存放着page解析出来的各种对象
                # 一般包括LTTextBox，LTFigure，LTImage，LTTextBoxHorizontal等等一些对像
                # 想要获取文本就得获取对象的text属性
                layout_x = int(layout.bbox[2])
                layout_y = int(layout.bbox[3])
                for x in layout:
                    # 读取所有文本
                    if isinstance(x, LTTextBoxHorizontal):
                        for line in x:
                            # 读取所有文本列
                            if isinstance(line, LTTextLine):
                                # 读取每一文本列
                                result = line.get_text().lower()
                                # 设置字符串匹配起始值
                                num_code = 0
                                # 匹配 原 字符串
                                # 判断要匹配的字符串code_source 中有几个匹配的值,循环几次
                                for i in range(result.count(code_source.lower())):
                                    # 关键字匹配，并获取在文本列的第几位
                                    codetwo = result.find(code_source, num_code)
                                    # 下次从这个值往后开始匹配
                                    num_code = codetwo + 1
                                    i = 0
                                    for char in line:
                                        # 读取每一个字符
                                        if isinstance(char, LTChar):
                                            # print(char.get_text())
                                            if i == codetwo:
                                                size_x.append(float(char.bbox[0]))
                                                size_y.append(float(char.bbox[1]))
                                                size_font.append(float(char.size))
                                        i += 1

            except e:
                print(e)
                print("The watermark file Create Failed")
                return 1
            # 根据匹配文本的坐标系，生成替换水印
            # 默认大小为
            name_without_postfix = os.path.splitext(pdf_file)[0]
            path_tmp = watermark_dir + "/" + name_without_postfix
            check_dir(path_tmp)
            mark = canvas.Canvas(path_tmp+"/" + str(pg) + ".pdf", pagesize=(layout_x, layout_y))

            #  生成每一页水印pdf ，将要代替的文字code_new ,写到匹配到的文字source_code位置，（用于合并覆盖)
            j = 0
            for i in size_x:
                # 移动坐标原点(坐标系左下为(0,0))
                x = float(i)
                y = float(size_y[j])

                # 指定描边的颜色
                mark.setStrokeColorRGB(1, 1, 1)
                # 指定填充颜色
                mark.setFillColorRGB(1, 1, 1)
                # 画一个矩形
                mark.rect(x, y, size_font[j] * 2.8, size_font[j] * 0.8, fill=1)
                j += 1
            z = 0
            for i in size_x:
                x = float(i)
                y = float(size_y[z])
                # 设置字体
                mark.setFont("Helvetica", size_font[z] * 0.8)

                # 指定填充颜色
                mark.setFillColorRGB(0, 0, 0)
                # 设置透明度，1为不透明
                mark.setFillAlpha(1)
                # 在水印pdf中写入一个字符，输入为全大写
                mark.drawString(x + (size_font[z] * 0.1), y + (size_font[z] * 0.125), code_new.upper())
                # 匹配是否为大文件，选择遮盖logo
                if size_font[z] > 20:
                    # 指定描边的颜色
                    mark.setStrokeColorRGB(1, 1, 1)
                    # 指定填充颜色
                    mark.setFillColorRGB(1, 1, 1)
                    # 画一个矩形
                    mark.rect(x - 35, y - 2, 35, size_font[z] * 0.9, fill=1)
                z += 1

            # 画一个空白矩阵
            mark.rect(0, 0, 0, 0, fill=1)
            # 关闭并保存pdf文件
            mark.save()
            pg += 1

        print('The watermark file Create Successful.')


def create_watermark_pdf_add(input_pdf):
    # 指定合拼PDF的源文件,输出到
    i_path_pdf = "./" + watermark_dir + "/" + input_pdf
    water_tmp_dir = "./" + watermark_dir + "/" + os.path.splitext(input_pdf)[0]
    pdf_reader = PdfFileReader(i_path_pdf)
    pdf_writer = PdfFileWriter()
    a = 0
    # 插入匹配成功生成的水印文件
    for page in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page)
        # 遍历pdf水印目录下的生成的水印文件
        watermark_obj = PdfFileReader(water_tmp_dir + "/" + str(a) + '.pdf')
        watermark_page = watermark_obj.getPage(0)
        page.mergePage(watermark_page)
        pdf_writer.addPage(page)
        a += 1
        # print(pdf_reader.getNumPages())
        # 获取完成的页数
        rate = float(a) / float(pdf_reader.getNumPages())
        # 输入完成的进度
        print('File conversion completed :' + '%.2f%%' % (rate * 100) + '.')

        # break
    # 存储到 {output} 目录，以原名称存储
    o_pdf_path = "./" + output_dir + "/" + input_pdf
    with open(o_pdf_path, 'wb') as out:
        # 设置密码
        # pdf_writer.encrypt('pega.1234')
        pdf_writer.write(out)
        # 输出完成信息
        print(input_pdf + "File second add watermark successful.")


# 处理pdf文件，将code_source 替换成 code_new
# 参数arg_list 是一个列表  【pdf，code_source,code_new】
def process_sop(arg_list):
    pdf = arg_list[0]
    code_s = arg_list[1]
    code_n = arg_list[2]
    copy_to_other_dir(pdf, watermark_dir)
    # parse(pdf, code_s, code_n)
    status = parse(pdf, code_s, code_n)
    if status == 1:
        # print('file:' + pdf + 'is protected and cannot be read.')
        return 2
    create_watermark_pdf_add(pdf)
    base_pdf = os.path.splitext(pdf)[0]
    path_rem = "./" + watermark_dir + "/" + base_pdf
    remove_dir(path_rem)
    file_path_rem = "./" + watermark_dir + "/" + pdf
    remove_file(file_path_rem)


# 并发
if __name__ == '__main__':
    freeze_support()
    code_source = input('请输入要替换的 原始 文字（Enter the source code that needs to be replaced）:')
    code_new = input('请输入要替换成的 新的 文字（Need to output the new code）:')
    check_dir(output_dir)
    check_dir(watermark_dir)
    pdf_info = get_pdf_namelist(read_dir)
    print(pdf_info[0])
    print(pdf_info[1])
    arg_list = []
    print("参数")
    arg_list = list(map(lambda x: [x, code_source, code_new], pdf_info[0]))
    print(arg_list)

    pool = multiprocessing.Pool(processes=3)  # 限制并行进程数为3
    pool.map(process_sop, arg_list)  # 创建进程池，调用函数a，传入参数为list,此参数必须是一个可迭代对象,因为map是在迭代创建每个进程
