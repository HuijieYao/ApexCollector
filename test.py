import cv2
import numpy as np

image = cv2.imdecode(
    np.fromfile("E:/sarom/Desktop/A026-5楼-2022年04月30日-晚上-抗原检测结果/501-3人/1951119_朱智舟_0430.png", dtype=np.uint8), 1)
h, w = image.shape[:2]  # 获取图像的高和宽
blured = cv2.blur(image, (5, 5))  # 对图像进行均值滤波（低通滤波）来去掉噪声
tmp = cv2.resize(blured, dsize=(450, 800), interpolation=cv2.INTER_CUBIC)
cv2.imshow("blured", tmp)
mask = np.zeros((h + 2, w + 2), np.uint8)  # 掩码长和宽都比输入图像多两个像素点，满水填充不会超出掩码的非零边缘
cv2.floodFill(blured, mask, (w - 1, h - 1), (255, 255, 255), (2, 2, 2), (3, 3, 3),
              8)  # 进行泛洪填充，类似PS中的魔棒工具，这里用于清除背景
gray = cv2.cvtColor(blured, cv2.COLOR_BGR2GRAY)  # 得到灰度图
tmp = cv2.resize(gray, dsize=(450, 800), interpolation=cv2.INTER_CUBIC)
cv2.imshow("gray", tmp)
kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (100, 150))  # 定义结构元素
opened = cv2.morphologyEx(gray, cv2.MORPH_OPEN, kernel)  # 开运算去除背景噪声
tmp = cv2.resize(opened, dsize=(450, 800), interpolation=cv2.INTER_CUBIC)
cv2.imshow("opened", tmp)
closed = cv2.morphologyEx(opened, cv2.MORPH_CLOSE, kernel)  # 闭运算填充目标内的孔洞
tmp = cv2.resize(tmp, dsize=(450, 800), interpolation=cv2.INTER_CUBIC)
cv2.imshow("closed", tmp)
ret, binary = cv2.threshold(closed, 250, 255, cv2.THRESH_BINARY)  # 二值化
tmp = cv2.resize(binary, dsize=(450, 800), interpolation=cv2.INTER_CUBIC)
cv2.imshow("binary", tmp)
contours, hierarchy = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)  # 找到轮廓
cv2.drawContours(image, contours, -1, (0, 0, 255), 10)  # 绘制轮廓
image = cv2.resize(image, dsize=(450, 800), interpolation=cv2.INTER_CUBIC)
cv2.imshow("result", image)
cv2.waitKey(0)
