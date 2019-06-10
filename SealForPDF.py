import os
import glob
import cv2  # 需要opencv的环境  通过anaconda直接安装
import numpy as np  # 需要安装numpy库
from PIL import Image
import fitz  # pdf转图片库
import re  # 正则匹配库
import pdfkit
import configparser  # 读取配置文件的库
import time
import math
import sys
from reportlab.lib.pagesizes import A4,landscape
from reportlab.pdfgen import canvas


# ------------------reportlab 绘图片方法 ---------------------
def reportDraw(Image):
    filename = 'test'
    f_png = filename + '.png'
    print(f_png)
    f_pdf = filename + '.pdf'
    (w,h) = landscape(A4)
    c = canvas.Canvas(f_pdf, pagesize=landscape(A4))
    c.drawImage(Image, 0, 0, w, h)
    c.save()
    print('DONE !!!!!')
# ---------------------先将 txt 转 HTML 然后再把 HTML 转为 pdf -------------------------
def txt2htm(txtName, htmlName):
    # txt = open(txtName.decode('utf-8'))
    txt = open(txtName, encoding='utf-8')  # 这里需要encoding = 'utf-8' 编码一下，因为Windows下默认是gbk编码   Mac默认是utf-8 就不用写encoding了

    filename = str(txtName).split('.')[0]
    # htmlName = filename + ".html"

    # htm = open(htmlName.decode('utf-8'), "w")
    htm = open(htmlName, "w", encoding='utf-8')
    htm.write('<html><head><title>test</title>')
    htm.write('<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head>')
    htm.write('<body><pre>')
    for line in txt:
        # htm.write('<b>'+str(line).decode('utf-8')+'</b>')
        htm.write('<b>' + str(line) + '</b>')

    htm.write('</pre></html>')
    #


    txt.close()
    htm.close()
    print('************pdf转换html成功************')

# -------------------------------------将 pdf 按页转为图片-----------------------------------
def run(target_file):
    #  打开PDF文件，生成一个对象
    doc = fitz.open(target_file)
    print('******当前项目路径:------>', os.getcwd())
    pic_path = os.path.join(os.getcwd(), 'all_pages')
    # pic_path = r'E:\PythonFormat\PDFfromConfig\all_pages'
    if not os.path.exists(pic_path):
        os.mkdir(pic_path)
    for pg in range(doc.pageCount):
        page = doc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为2，这将为我们生成分辨率提高四倍的图像。
        zoom_x = 3
        zoom_y = 3
        trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pm = page.getPixmap(matrix=trans, alpha=False)
        # pm.writePNG('%s.png' % pg)
        temppic_path = os.path.join(pic_path, (str(pg) + '.png'))
        pm.writePNG(temppic_path)

    handlePageNum = str(doc.pageCount)
    doc.close()
    print('************pdf转图片完成************')
    return handlePageNum


# -----------------------------------------将要添加的印章图片和目标图片进行叠加------------------------------------
# 根据背景图的宽高逆推印章的位置 从而进行合成
def ImageMixed(bg_image, target_image, bg_width, bg_height, point_x, point_y, pages, page):
    img = cv2.imread(bg_image)
    co = cv2.imread(target_image, -1)
    scr_channels = cv2.split(co)
    dstt_channels = cv2.split(img)
    b, g, r, a = cv2.split(co)
    if point_x+154>bg_width or point_y+154>bg_height:
        print('图片超出pdf范围！！！')
    else:
        for i in range(3):
            # 由于坐标系是向下为X  向右为Y
            dstt_channels[i][point_y:point_y+154, point_x:point_x+154] = dstt_channels[i][point_y:point_y+154, point_x:point_x+154] * (255.0 - a) / 255
            dstt_channels[i][point_y:point_y+154, point_x:point_x+154] += np.array(scr_channels[i] * (a / 255), dtype=np.uint8)
        # cv2.imwrite("img_target.png", cv2.merge(dstt_channels))
        cv2.imwrite(bg_image, cv2.merge(dstt_channels))

    print('共{}页************第{}页图片叠加完成************'.format(pages, page))

# --------------------------------------将图片转成 pdf ----------------------------------------
def pic2pdf(Source_pdf_path):
    doc = fitz.open()
    for img in sorted(glob.glob('all_pages/*')):
        imgdoc = fitz.open(img)
        pdfbytes = imgdoc.convertToPDF()
        imgpdf = fitz.open('pdf', pdfbytes)
        doc.insertPDF(imgpdf)
    if os.path.exists(Source_pdf_path):
        os.remove(Source_pdf_path)
    doc.save(Source_pdf_path)
    doc.close()
    print('************图片再转pdf完成************\n')

# --------------------------递归删除临时存放单独文件的图片目录--------------------
# 递归删除临时存放pdf每页图片的文件夹
def rmDir(all_page_path):
    # 获取目录下的所有文件目录以列表形式返回
    pathList = os.listdir(all_page_path)
    # 迭代获取每一个文件或目录
    for file in pathList:
        # 拼凑为完整路径
        newPath = os.path.join(all_page_path, file)
        # 判断是否为文件
        if os.path.isfile(newPath):
            # 删除文件
            os.remove(newPath)
        # 判断是否为目录
        if os.path.isdir(newPath):
            # 递归调用自己
            rmDir(newPath)
    # 把目录中的文件都删除后则删除文件夹
    os.rmdir(all_page_path)



def main():

    # 拼接印章路径
    image_path = os.path.join(os.getcwd(), 'Seal.png')

    # 拼接配置文件路径
    config_path = os.path.join(os.getcwd(), 'Sett_Bill_config.ini')

    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')
    #                         所有关于文件路劲的获取
    # 拿到所有资源的一个路径
    allSettlement_Dir = config.get(section='Dir', option='Settlement_Source_dir')
    # 拿到源操作文件的txt的格式 Settlement_Format
    Settlement_Format = config.get(section='Dir', option='Settlement_Format')
    Settlement_Format_html = config.get(section='Dir', option='Settlement_Format_html')
    Settlement_Format_pdf = config.get(section='Dir', option='Settlement_Format_pdf')
    Settlement_SealPDF_Dir = config.get(section='Dir', option='Settlement_SealPDF_dir')
    # 拿到源操作文件的电子命名格式 Settlement_E_Bill_Format
    # Settlement_E_Bill_Format = config.get(section='Dir', option='Settlement_E_Bill_Format')

    # 关于印章位置的配置获取
    Point_x = config.get(section='SealPoint', option='Point_X')
    Point_y = config.get(section='SealPoint', option='Point_Y')


    # 拿到Fundaccount
    Fundaccount = config.get(section='JS', option='Fundaccount_Filter')
    Fundaccount_list = Fundaccount.split('|')

    Fundaccount_len = len(Fundaccount_list)
    # 获取当前日期格式为 YYYYMMDD
    now_day = time.strftime('%Y%m%d', time.localtime())
    for i in range(Fundaccount_len):

        Settlement_Format_list = Settlement_Format.split('_')
        Settlement_Format_html_list = Settlement_Format_html.split('_')
        Settlement_Format_pdf_list = Settlement_Format_pdf.split('_')
        # print(Settlement_Format_list)
        # print(Settlement_Format_html_list)
        # print(Settlement_Format_pdf_list)

        # 开始拼接源txt路径
        #   先拼接当前txt 的name
        Source_txt_name = now_day + '_' + Fundaccount_list[i] + '_' + Settlement_Format_list[2]
        Source_html_name = now_day + '_' + Fundaccount_list[i] + '_' + Settlement_Format_html_list[2]
        Source_pdf_name = now_day + '_' + Fundaccount_list[i] + '_' + Settlement_Format_pdf_list[2]
        # print(Source_txt_name)
        #   再拼接文件的os路径
        Source_txt_path = os.path.join(allSettlement_Dir, Source_txt_name)
        Source_html_path = os.path.join(Settlement_SealPDF_Dir, Source_html_name)
        Source_pdf_path = os.path.join(Settlement_SealPDF_Dir, Source_pdf_name)
        # print(Source_path)

        if not os.path.exists(Source_txt_path):
            print('配置文件中TXT文件路径不存在\n请修改配置文件中的Settlement_Source_dir字段为自己存放账单txt的路径')
        else:
            txt2htm(Source_txt_path, Source_html_path)

        pdfkit.from_file(Source_html_path, str(Source_pdf_path))



        pagecount = run(Source_pdf_path)
        pages = int(pagecount)
        # pdf每页图片文件夹的路径拼接
        path = os.path.join(os.getcwd(), "all_pages")


        for i in range(pages):
            # 每个图片的路径拼接 （这里的pagecount目前只是指定第一页还是第几页  单页的盖章模式）
            pic_path = os.path.join(path, (str(i) + '.png'))
            # 这里打开目标背景图 并获取其size， 方便对其边界进行判断
            fp = open(pic_path, 'rb')
            img = Image.open(fp)
            fp.close()

            bg_width = img.size[0]  # 每页的pdf图片的宽
            bg_heigh = img.size[1]  # 每页的pdf图片的高
            point_x = math.ceil(bg_width*eval(Point_x))
            point_y = math.ceil(bg_heigh*eval(Point_y))
            # print('point_x{} point_y{}'.format(point_x, point_y))
            # print(pic_path, bg_width, bg_heigh)

            # pdf图片和印章图片的合成
            # 根据背景图的宽高逆推印章的位置 从而进行合成
            ImageMixed(pic_path, image_path, bg_width, bg_heigh, point_x, point_y, pages, i+1)

        # 调用图片转pdf
        pic2pdf(Source_pdf_path)

        # 删除中间产生的html文件
        os.remove(Source_html_path)

        # 调用递归删除目录
        all_page_path = 'all_pages'
        rmDir(all_page_path)



main()
input('按任意键退出')
