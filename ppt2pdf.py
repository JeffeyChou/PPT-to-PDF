#!python3
'''
Automatically convert a PPT/PPTX to PNG files and combine them to a PDF after applying watermark.

'''

import os
import math
import win32com.client
import shutil
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
__WaterMarkDir = 'watermark.txt'
__WaterMarkInline = "test"
__if_watermark = 2
# 0 for no watermark,
# 1 for inline watermark,
# 2 for using text in "watermark.txt" as watermark


def ppt2png(filename, dst_filename):
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # ppt.DisplayAlerts = False
    pptSel = ppt.Presentations.Open(filename, WithWindow=False)
    pptSel.SaveAs(dst_filename, 18)  # with 17, jpeg
    ppt.Quit()


def AddWaterMark(imgFile, textMark):
    img = Image.open(imgFile)
    imgWidth, imgHeight = img.size
    # http://blog.csdn.net/Dou_CO/article/details/17715919
    textImgW = int(imgWidth * 1.5)  # 确定写文字图片的尺寸，要比照片大
    textImgH = int(imgHeight * 1.5)
    blank = Image.new("RGB", (textImgW, textImgH), "white")  # 创建用于添加文字的空白图像
    d = ImageDraw.Draw(blank)
    d.ink = 0 + 0 * 256 + 0 * 256 * 256
    markFont = ImageFont.truetype('simhei.ttf', size=180)
    fontWidth, fontHeight = markFont.getsize(textMark)
    d.text(((textImgW - fontWidth)/2, (textImgH - fontHeight)/2),
           textMark, font=markFont)
    textRotate = blank.rotate(30)

    rLen = math.sqrt((fontWidth/2)**2+(fontHeight/2)**2)
    oriAngle = math.atan(fontHeight/fontWidth)
    cropW = rLen*math.cos(oriAngle + math.pi/6) * 4  # 被截取区域的宽高
    cropH = rLen*math.sin(oriAngle + math.pi/6) * 4
    box = [int((textImgW-cropW)/2-1), int((textImgH-cropH)/2-1)-50,
           int((textImgW+cropW)/2+1), int((textImgH+cropH)/2+1)]
    textImg = textRotate.crop(box)  # 截取文字图片
    pasteW, pasteH = textImg.size
    # 旋转后的文字图片粘贴在一个新的blank图像上
    textBlank = Image.new("RGB", (imgWidth, imgHeight), "white")
    pasteBox = (int((imgWidth-pasteW)/2-1), int((imgHeight-pasteH)/2-1))
    textBlank.paste(textImg, pasteBox)
    waterImage = Image.blend(img.convert('RGB'), textBlank, 0.1)

    fileDir = os.path.dirname(imgFile)
    fileName = os.path.join(fileDir, os.path.basename(imgFile))
    waterImage.save(fileName, 'png')


def pic2pdf(path, recursion=None, pictureType=None, sizeMode=None, width=None, height=None, fit=None, save=None):
    """
    Parameters
    ----------
    path : string
               path of the pictures
    pictureType : list
                              type of pictures,for example :jpg,png...
    sizeMode : int
               None or 0 for pdf's pagesize is the biggest of all the pictures
               1 for pdf's pagesize is the min of all the pictures
               2 for pdf's pagesize is the given value of width and height
               to choose how to determine the size of pdf
    width : int
                    width of the pdf page
    height : int
                    height of the pdf page
    fit : boolean
               None or False for fit the picture size to pagesize
               True for keep the size of the pictures
               wether to keep the picture size or not
    save : string
               path to save the pdf
    """

    filelist = os.listdir(path)
    filelist = [os.path.join(path, f) for f in filelist]
    filelist.sort(key=lambda x: os.path.getmtime(x))

    maxw = 0
    maxh = 0
    if not sizeMode:
        for i in filelist:
            # print('----'+i)
            im = Image.open(i)
            if maxw < im.size[0]:
                maxw = im.size[0]
            if maxh < im.size[1]:
                maxh = im.size[1]
    elif sizeMode == 1:
        maxw = 999999
        maxh = 999999
        for i in filelist:
            im = Image.open(i)
            if maxw > im.size[0]:
                maxw = im.size[0]
            if maxh > im.size[1]:
                maxh = im.size[1]
    else:
        if not width or height:
            raise Exception("no width or height provid")
        maxw = width
        maxh = height

    maxsize = (maxw, maxh)
    if not save:
        filename_pdf = os.path.join(path, path.split('\\')[-1])
    else:
        filename_pdf = os.path.join(save, path.split('\\')[-1])

    filename_pdf = filename_pdf + '.pdf'
    print('Ready to bulit' + filename_pdf)
    c = canvas.Canvas(filename_pdf, pagesize=maxsize)

    lenoflist = len(filelist)
    for i in range(lenoflist):
        print(filelist[i])
        (w, h) = maxsize
        if fit:
            c.drawImage(filelist[i], 0, 0)
        else:
            c.drawImage(filelist[i], 0, 0, maxw, maxh)
        c.showPage()
    c.save()


def main(__if_watermark=0):
    ppt_dir = os.getcwd()
    WatermarkDir = os.path.join(ppt_dir, __WaterMarkDir)
    WatermarkCon = open(WatermarkDir, encoding="utf8")
    for fn in (FileNames for FileNames in os.listdir(ppt_dir) if FileNames.endswith(('.ppt', '.pptx'))):
        PptName = os.path.splitext(fn)[0]
        print(PptName)
        ppt_file = os.path.join(ppt_dir, fn)
        img_file = os.path.join(ppt_dir, PptName+'.png')
        ppt2png(ppt_file, img_file)
        img_dir = os.path.join(ppt_dir, PptName)
        imgFileList = os.listdir(img_dir)
        imgFileList = [os.path.join(img_dir, f) for f in imgFileList]
        imgFileList.sort(key=lambda x: os.path.getmtime(x))
        if __if_watermark:
            if __if_watermark == 2:
                for markText in WatermarkCon.readlines():
                    os.makedirs(img_dir, exist_ok=True)
                    for imgFile in imgFileList:
                        AddWaterMark(imgFile, markText.strip('\n'))
                    print(f'Complete adding watermark in {PptName}!')
            else:
                os.makedirs(img_dir, exist_ok=True)
                for imgFile in imgFileList:
                    AddWaterMark(imgFile, __WaterMarkInline)
                print(f'Complete adding watermark in {PptName}!')
        pic2pdf(path=img_dir, save=ppt_dir)
        print(f"Transfer {PptName} to PDF successfully!")
        shutil.rmtree(img_dir)


if __name__ == "__main__":
    main(__if_watermark)
