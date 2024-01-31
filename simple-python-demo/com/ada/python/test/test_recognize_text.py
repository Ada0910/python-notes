# import pytesseract
# from PIL import Image
# import numpy as np
# 
# # 截取图片
# screenshot_path = "E:\\NC\\picture\\screenshot.png"
# 
# # 加载验证码图片
# image = Image.open(screenshot_path)
# 
# # 预处理图像（根据实际情况进行调整）
# # 进行图像二值化、降噪等处理
# pytesseract.pytesseract.tesseract_cmd = r'E:\\Tesseract-ocr\\tesseract.exe'
# # 图片灰度处理
# imageObject = np.array(image.convert('L'), 'f')
# # 使用Pytesseract进行字符识别
# text = pytesseract.image_to_string(image)
# 
# # 输出识别结果
# print(text)

import cv2 as cv
import pytesseract
from PIL import Image


import cv2 as cv
import pytesseract
from PIL import Image


def recognize_text(image):
    # 边缘保留滤波  去噪
    blur =cv.pyrMeanShiftFiltering(image, sp=8, sr=60)
    cv.imshow('dst', blur)
    # 灰度图像
    gray = cv.cvtColor(blur, cv.COLOR_BGR2GRAY)
    # 二值化
    ret, binary = cv.threshold(gray, 0, 255, cv.THRESH_BINARY_INV | cv.THRESH_OTSU)
    print(f'二值化自适应阈值：{ret}')
    cv.imshow('binary', binary)
    # 形态学操作  获取结构元素  开操作
    kernel = cv.getStructuringElement(cv.MORPH_RECT, (3, 2))
    bin1 = cv.morphologyEx(binary, cv.MORPH_OPEN, kernel)
    cv.imshow('bin1', bin1)
    kernel = cv.getStructuringElement(cv.MORPH_OPEN, (2, 3))
    bin2 = cv.morphologyEx(bin1, cv.MORPH_OPEN, kernel)
    cv.imshow('bin2', bin2)
    # 逻辑运算  让背景为白色  字体为黑  便于识别
    cv.bitwise_not(bin2, bin2)
    cv.imshow('binary-image', bin2)
    # 识别
    test_message = Image.fromarray(bin2)
    pytesseract.pytesseract.tesseract_cmd = r'E:\\Tesseract-ocr\\tesseract.exe'
    text = pytesseract.image_to_string(test_message)
    print(f'识别结果：{text}')


src = cv.imread(r'E:\\NC\\picture\\screenshot.png')
cv.imshow('input image', src)
recognize_text(src)
cv.waitKey(0)
cv.destroyAllWindows()

# import cv2
# import pytesseract
# 
# # 加载验证码图片
# image = cv2.imread('E:\\NC\\picture\\screenshot.png')
# 
# # 将图像转换为灰度
# gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
# 
# # 对灰度图像进行二值化处理
# _, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
# 
# pytesseract.pytesseract.tesseract_cmd = r'E:\\Tesseract-ocr\\tesseract.exe'
# # 使用 Pytesseract 进行字符识别
# text = pytesseract.image_to_string(binary)
# 
# # 输出识别结果
# print(text)