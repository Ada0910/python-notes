"""
Python 获取由java元素边编写的界面的所有复选框的数量
"""
import cv2
import pyautogui

# 加载界面截图和复选框模板图像
screenshot = cv2.imread('E:\\NC\\picture\\origin\\screenshot.jpg', 0)
template = cv2.imread('E:\\NC\\picture\\checkbox_template.jpg', 0)

# 获取模板图像的宽度和高度
template_height, template_width = template.shape

# 使用模板匹配方法进行匹配
result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)

# 设置匹配阈值
threshold = 0.8

# 获取匹配结果中大于阈值的位置
locations = []
for y in range(result.shape[0]):
    for x in range(result.shape[1]):
        if result[y, x] >= threshold:
            locations.append((x, y))

# 绘制矩形框标记复选框区域
for loc in locations:
    start_x = loc[0]
    start_y = loc[1]
    end_x = start_x + template_width
    end_y = start_y + template_height
    # 单击
    pyautogui.click((start_x+end_x)/2, (start_y+end_y)/2)
    cv2.rectangle(screenshot, (start_x, start_y), (end_x, end_y), (0, 255, 0), 2)

# 输出复选框数量
num_checkboxes = len(locations)
print("复选框数量:", num_checkboxes)

# 显示结果图像
cv2.imshow("Result", screenshot)
cv2.waitKey(0)
cv2.destroyAllWindows()


"""
Python 获取由java元素边编写的界面的所有复选框的数量并单击
这个是未经验证的
"""

import pyautogui
import cv2

# 加载截图
screenshot = cv2.imread("E:\\NC\\picture\\origin\\screenshot.jpg", 0)

# 加载复选框的图片（截取界面中的一个复选框作为模板）
checkbox_image = cv2.imread("E:\\NC\\picture\\checkbox_template.jpg", 0)  # 灰度图像

# 使用模板匹配找到复选框的位置
result = cv2.matchTemplate(screenshot, checkbox_image, cv2.TM_CCOEFF_NORMED)
locations = pyautogui.locateAllOnScreen(result, confidence=0.8)

# 选中所有的复选框
for loc in locations:
    checkbox_center = pyautogui.center(loc)
    pyautogui.click(checkbox_center)

# 返回复选框的数量
checkbox_count = len(locations)
print("复选框的数量：", checkbox_count)
