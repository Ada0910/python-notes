"""
桌面应用程序操作工具类
"""
import os
import time

import cv2
import keyboard
import pyautogui
import pyperclip
from PIL import ImageGrab
from pynput import mouse

"""
 查找图片并进行点击

Args:
    original_path：截图保存路径
    target_path： 目标图片保存路径
    filename: 目标图片文件名
    action : single: 单击 , double: 双击，right:鼠标右击
    extra_x: （选填）左移或右移
    extra_y：（选填）上移或者下移
Returns:
    无返回

Raises:
    异常情况的说明（如果有的话）
"""


def find_picture_click(original_path, target_path, filename, action, extra_x=0, extra_y=0):
    # 获取处理后的路径
    result = handle_path_file_name(original_path, target_path, filename)
    original_path_name, target_path_name = result

    # 获取目标图片左上角、右下角的坐标
    top_left, bottom_right = match_template(original_path_name, target_path_name)

    # 获取鼠标即将点击的x,y坐标
    x, y = get_click_point(top_left, bottom_right)

    if action == "single":
        # 单击鼠标左键
        # 参数1：移动x坐标到指定位置
        # 参数2：移动y坐标到指定位置
        # 参数3：点击次数
        # 无参数：在当前坐标单击
        pyautogui.click(x + extra_x, y + extra_y)
    elif action == "double":
        # 双击鼠标左键
        # 参数1：移动x坐标到指定位置
        # 参数2：移动y坐标到指定位置
        # 无参数：在当前坐标双击
        pyautogui.doubleClick(x + extra_x, y + extra_y)
    elif action == "right":
        # 点击鼠标右键
        # 参数1：移动x坐标到指定位置
        # 参数2：移动y坐标到指定位置
        # 无参数：在当前坐标点击
        pyautogui.rightClick(x + extra_x, y + extra_y)


"""
处理文件名，并保存图片

Args:
    original_path：截图保存路径
    target_path： 目标图片保存路径
    filename: 目标图片文件名
Returns:
    处理后的original_path_name全路径
    处理后的target_path_name全路径

Raises:
    异常情况的说明（如果有的话）
"""


def handle_path_file_name(original_path, target_path, filename):
    original_path = init_path(original_path)
    target_path = init_path(target_path)
    # 找到最后一个圆点的索引位置
    last_dot_index = filename.rfind(".")
    if last_dot_index != -1:
        file_name = filename[:last_dot_index]  # 获取文件名部分
        file_extension = filename[last_dot_index + 1:]  # 获取文件名后缀部分
    else:
        print("无法解析文件名和文件名后缀")
    # 截图的保存路径的绝对路径名
    original_path_name = original_path + file_name + "_origin." + file_extension
    # 目标图片的保存路径的绝对路径名
    target_path_name = target_path + file_name + "." + file_extension

    # 截图
    pic = ImageGrab.grab()

    # 保存图片
    pic.save(original_path_name)
    return original_path_name, target_path_name


"""
处理路径

Args:
    path : 需要处理的路径
    is_append_string : True追加末尾分隔符，False不逐年末位分隔符
Returns:
    处理后的路径
"""


def treat_path(path, is_append_string=False):
    if path.endswith("\\"):
        return path
    else:
        if is_append_string is True:
            return path + os.sep
        else:
            return path

    pass


"""
初始化路径，没有则创建，兼容末尾有斜杠和没斜杠

Args:
    target_path : 需要处理的路径
Returns:
    处理后的路径
"""


def init_path(target_path):
    result_path = treat_path(target_path, True)
    os.makedirs(result_path, exist_ok=True)
    return result_path


"""
 查找image_path图中，template_path目标图像的坐标

Args:
    image_path：截图保存路径
    template_path： 目标图片保存路径
Returns:
    返回目标图片在截图中的左上角、右下角的坐标

Raises:
    异常情况的说明（如果有的话）
"""


def match_template(image_path, template_path):
    # 读取图像
    image = cv2.imread(image_path)
    template = cv2.imread(template_path)
    # 返回匹配的结果
    result = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    top_left = max_loc
    bottom_right = (top_left[0] + template.shape[1], top_left[1] + template.shape[0])

    # 4.画矩形
    cv2.rectangle(image, top_left, bottom_right, (255, 0, 0), 5)
    cv2.imwrite(image_path, image)
    m = mouse.Controller()
    m.position = top_left
    # 返回左上角、右下角的坐标
    return top_left, bottom_right


"""
# 计算出鼠标要点击的xy坐标

Args:
    top_left：左上角坐标
    bottom_right： 右下角坐标
Returns:
   计算出鼠标要点击的xy坐标

Raises:
    异常情况的说明（如果有的话）
"""


def get_click_point(top_left, bottom_right):
    # 鼠标左键点击的坐标
    x = (top_left[0] + bottom_right[0]) / 2
    y = (top_left[1] + bottom_right[1]) / 2
    return x, y


"""
获取目标图片的中心坐标数组

Args:
    original_path：截图保存路径
    target_path： 目标图片保存路径
    filename: 目标图片文件名
Returns:
    获取匹配的目标图片的中心坐标数组

Raises:
    异常情况的说明（如果有的话）
"""


def get_target_locations(original_path, target_path, filename):
    temp = handle_path_file_name(original_path, target_path, filename)
    original_path_name, target_path_name = temp
    # 加载界面截图和复选框模板图像
    screenshot = cv2.imread(original_path_name, 0)
    template = cv2.imread(target_path_name, 0)

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

    target_locations = []
    for loc in locations:
        start_x = loc[0]
        start_y = loc[1]
        end_x = start_x + template_width
        end_y = start_y + template_height
        middle_x = (start_x + end_x) / 2
        middle_y = (start_y + end_y) / 2
        target_locations.append((middle_x, middle_y))
    return target_locations


"""
# 查找输入框并输入字符串

Args:
    original_path：截图保存路径
    target_path： 目标图片保存路径
    target_name: 目标图片文件名
    input_word： 输入字符串
Returns:
    无

Raises:
    异常情况的说明（如果有的话）
"""


def find_picture_click_and_input(orginal_path, target_path, target_name, input_word):
    find_picture_click(orginal_path, target_path, target_name, "double")
    # 输入字符串
    pyautogui.typewrite("", interval=0.1)
    pyautogui.typewrite(input_word, interval=0.1)
    keyboard.press_and_release("enter")


"""获取剪贴板数据"""


def clipboard_get():
    time.sleep(0.2)
    data = pyperclip.paste()  # 主要这里差别
    return data


"""
# 检查字符串是否有匹配的

Args:
    string：目标字符串
    value: 查找匹配字符串
Returns:
    True匹配；False不匹配

Raises:
    异常情况的说明（如果有的话）

"""


def check_string(string, value):
    if string.find(value) != -1:
        return True
    else:
        return False


