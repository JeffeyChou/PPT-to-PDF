from pdf2image import convert_from_path
import os
import math
from PIL import Image, ImageDraw, ImageFont
__WaterMarkDir = 'watermark.txt'
__WaterMarkInline = "test"
__if_watermark = 0


def pdf2png(PDF_file, png_dir):
    pages = convert_from_path(PDF_file, 500)
    os.makedirs(png_dir, exist_ok=True)
    for _ in range(len(pages)):
        prefix = str(_ + 1)
        filename = os.path.join(png_dir, prefix + '.png')
        pages[_].save(filename, 'PNG')


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


def main(__if_watermark=0):
    PDF_dir = os.getcwd()
    WatermarkDir = os.path.join(PDF_dir, __WaterMarkDir)
    WatermarkCon = open(WatermarkDir, encoding="utf8")
    for fn in (FileNames for FileNames in os.listdir(PDF_dir) if FileNames.endswith(('.pdf'))):
        PDF_name = os.path.splitext(fn)[0]
        print(PDF_name)
        PDF_file = os.path.join(PDF_dir, fn)
        png_dir = os.path.join(PDF_dir, PDF_name)
        pdf2png(PDF_file, png_dir)
        if __if_watermark:
            imgFileList = os.listdir(png_dir)
            imgFileList = [os.path.join(png_dir, f) for f in imgFileList]
            imgFileList.sort(key=lambda x: os.path.getmtime(x))
            if __if_watermark == 2:
                for markText in WatermarkCon.readlines():
                    for imgFile in imgFileList:
                        AddWaterMark(imgFile, markText.strip('\n'))
                    print(f'Complete adding watermark in {PDF_name}!')
            else:
                for imgFile in imgFileList:
                    AddWaterMark(imgFile, __WaterMarkInline)
                print(f'Complete adding watermark in {PDF_name}!')
        print(f"Transfer {PDF_name} to PNG successfully!")


if __name__ == "__main__":
    main(__if_watermark)
