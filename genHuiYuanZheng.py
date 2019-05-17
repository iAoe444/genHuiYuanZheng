import qrcode
import os
from pystrich.code128 import Code128Encoder
import cv2
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import xlrd

import numpy as np

# create at 2019/05/16 22:20

nameFont = ImageFont.truetype("NotoSansHans-Bold.otf",60)
yearFont = ImageFont.truetype("NotoSansHans-Bold.otf",45)
# 打开工作簿
workbook = xlrd.open_workbook('人员信息.xlsx')
# 获取工作表
table = workbook.sheet_by_index(0)
year = int(table.cell(0,0).value)

# 生成二维码
def genQrCode(msg):
    qr = qrcode.QRCode(
        version=1,      # 用于控制二维码的格子的数量，最小值为1，12*12
        error_correction=qrcode.constants.ERROR_CORRECT_M,     # 用于控制错误纠错率
        box_size=10,    # 用于控制每个二维码格子的大小
        border=1,       # 用于控制外层白色边框的大小
    )
    qr.add_data(msg)
    qr.make(fit=True)
    img = qr.make_image()

    img.save('qrCode/'+msg+'.png')

    # 对二维码进行缩放，缩放到157*157
    reSizeImg('qrCode/'+msg+'.png',width=157)
    return 'qrCode/'+msg+'.png'

# 生成条形码
def genBarCode(msg):
    options = {
        'height':120,        # 条形码的高度
        'label_border':10,   # 标签和二维码之间的距离
    }
    encoder = Code128Encoder(msg,options=options)
    encoder.save('barCode/'+msg+'.png')

    # 对条形码进行裁剪
    img = cv2.imread('barCode/'+msg+'.png')
    cropped = img[2:105,0:460]
    cv2.imwrite('barCode/'+msg+'.png',cropped)

    # 对条形码进行缩放，缩放到558宽度
    reSizeImg('barCode/'+msg+'.png',width=558)

    return 'barCode/'+msg+'.png'

# 创建文件夹
def makeDirs(*dirs):
    for dir in dirs:
        # 判断是否存在文件夹，不存在就创建该文件夹
        isExists = os.path.exists(dir)
        if not isExists:
            os.makedirs(dir)

# 图片缩放
def reSizeImg(imgPath,width=0,height=0):
    img=cv2.imdecode(np.fromfile(imgPath,dtype=np.uint8),-1)    # 图片路径可能带中文，用这个方法读取图片
    imgHeight, imgWidth = img.shape[:2]   # 获取图片的高和宽
    # 如果wdith=0,意思是按高度不变，等比例缩放，所以需要计算width
    if width==0:
        width = int(imgWidth * height / imgHeight)
    # 如果height=0,意思是按宽度不变，等比例缩放，所以需要计算height
    if height==0:
        height = int(imgHeight * width / imgWidth)
    reSizedImg = cv2.resize(img,(width,height),interpolation=cv2.INTER_CUBIC)
    # 保存图片为原来的路径
    cv2.imwrite(imgPath,reSizedImg)

# 贴图操作，这里选用PIL库，较为简单
def pasteImg(pastedImgPath,imgPath,outPutPath,x=0,y=0):
    pastedImg = Image.open(pastedImgPath)
    img = Image.open(imgPath)
    # 如果图片有透明度通道，那么需要提取出来
    if len(img.split())==4:
        r,g,b,a = img.split()
        mask = a
    else:
        mask = None
    pastedImg.paste(img,(x,y),mask)   # 如果图片有透明度，需要把透明度传过去
    pastedImg.save(outPutPath)

# TODO 目前切割边缘不理想，需优化圆形切割算法
# 切割圆形图片
def genCicleImg():
    ima = Image.open("1.jpg").convert("RGBA") 
    size = ima.size 
    r2 = min(size[0], size[1]) 
    # 最后生成圆的半径 
    r3 = float(r2/2)
    imb = Image.new('RGBA', (int(r3*2), int(r3*2)),(255,255,255,0)) 
    pima = ima.load() # 像素的访问对象 
    pimb = imb.load() 
    r = float(r2/2) #圆心横坐标 
    for i in range(r2): 
        for j in range(r2):  
            lx = abs(i-r) #到圆心距离的横坐标 
            ly = abs(j-r)#到圆心距离的纵坐标 
            l = (pow(lx,2) + pow(ly,2))** 0.5 # 三角函数 半径 
            if l < r3: 
                pimb[i-(r-r3),j-(r-r3)] = pima[i,j]
    imb.save("head/1.png")
    return "head/1.png"

# TODO 文字居中问题
# 图片添加文字
def addText(imgPath,text,font,x=0,y=0):
    img = Image.open(imgPath)
    # 添加文字
    draw = ImageDraw.Draw(img)
    draw.text((x,y),text,(255,255,255),font=font)
    draw = ImageDraw.Draw(img)
    img.save(imgPath)

# 生成会员证
# TODO 生成图片分辨率问题，图片分辨率要达到300dpi
def genHuiYuanZheng(num,name):
    # -----------添加正面二维码----------------
    # 生成二维码
    qrImgPath = genQrCode(num)
    # 把二维码，贴到上面
    pasteImg('正面.png',qrImgPath,'front/'+num+'.png',632,82)

    # -------------添加头像--------------------
    # 寻找是否是会员图片
    if os.path.exists("相片/"+name+".png"):
        headImgPath="相片/"+name+".png"
        # 预处理一下图片，先进行缩放操作，如果事先的图片是312*312，则不需要这步操作
        reSizeImg("相片/"+name+".png",312,312)
    else:
        headImgPath="会徽.png"
    # 由于生成圆形头像的图片还有缺陷，故暂时用这种方式解决圆形头像的问题
    pasteImg(headImgPath,'圆形头像.png','head/'+num+'.png',0,0)
    # 把头像贴到上面
    pasteImg('front/'+num+'.png','head/'+num+'.png','front/'+num+'.png',267,400)

    # -------------添加文字---------------------
    text = ' '.join(name)
    if len(name)==2:
        addText('front/'+num+'.png',text,nameFont,358,755)
    elif len(name)==3:
        addText('front/'+num+'.png',text,nameFont,318,755)
    elif len(name)==4:
        addText('front/'+num+'.png',text,nameFont,280,755)

    # ------------添加背面条形码----------------
    # 生成条形码
    barImgPath = genBarCode(num)
    # 把二维码，贴到上面
    pasteImg('背面.png',barImgPath,'back/'+num+'.png',141,980)
    addText('back/'+num+'.png',str(year),yearFont,684,95)

if __name__ == "__main__":
    makeDirs('front','back','head','qrCode','barCode')  # 预先生成文件夹
    for i in range(2,table.nrows):  # 读取excel里面的内容
        name = table.cell(i,0).value
        num = table.cell(i,1).value
        genHuiYuanZheng(num,name)
        print(num+name+" 生成完毕")