import time
import subprocess
import pygetwindow as gw
import cv2
import pyautogui
import keyboard
import os
from pynput import mouse
from PIL import ImageGrab


# ------------------------------方法区------------------------------

# 初始化路径，没有则创建，兼容末尾有斜杠和没斜杠
def init_path(target_path):
    result_path = treat_path(target_path, True)
    os.makedirs(result_path, exist_ok=True)
    return result_path


# 处理路径
# param path: 路径信息
# param is_append_string:True追加末尾分隔符False不逐年末位分隔符
# return:处理后的路径
def treat_path(path, is_append_string=False):
    if path.endswith("\\"):
        return path
    else:
        if is_append_string is True:
            return path + os.sep
        else:
            return path

    pass


# 获取窗口并最大化
def get_window():
    # 获取指定窗口并置顶，最大化窗口
    window = gw.getWindowsWithTitle('Yonyou UClient')[0]
    window.setAlwaysOnTop(True)
    window.maximize()


# 查找image_path图中，template_path目标图像的坐标 
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


# 计算出鼠标要点击的xy坐标
def get_click_point(top_left, bottom_right):
    # 鼠标左键点击的坐标
    x = (top_left[0] + bottom_right[0]) / 2
    y = (top_left[1] + bottom_right[1]) / 2
    return x, y


# 查找图片并点击
# action : single: 单击 , double: 双击
def find_picture_click(orginal_path, target_path, original_name, target_name, action):
    # 截图的保存路径的绝对路径名
    orginal_path_name = orginal_path + original_name
    # 目标图片的保存路径的绝对路径名
    target_path_name = target_path + target_name

    # 截图
    pic = ImageGrab.grab()

    # 保存图片
    pic.save(orginal_path_name)

    # 获取目标图片左上角、右下角的坐标
    top_left, bottom_right = match_template(orginal_path_name, target_path_name)
    # print("左上角、右下角的坐标分别是：", (top_left, bottom_right))

    # 获取鼠标即将点击的x,y坐标
    x, y = get_click_point(top_left, bottom_right)
    # print("鼠标点击的[x,y]坐标是：", (x, y))

    if action == "single":
        # 单击鼠标左键
        # 参数1：移动x坐标到指定位置
        # 参数2：移动y坐标到指定位置
        # 参数3：点击次数
        # 无参数：在当前坐标单击
        pyautogui.click(x, y)
    elif action == "double":
        # 双击鼠标左键
        # 参数1：移动x坐标到指定位置
        # 参数2：移动y坐标到指定位置
        # 无参数：在当前坐标双击
        pyautogui.doubleClick(x, y)


# 查找输入框并输入字符串
def find_picture_click_and_input(orginal_path, target_path, original_name, target_name, input_word):
    find_picture_click(orginal_path, target_path, original_name, target_name, "double")
    # 输入字符串
    pyautogui.typewrite("", interval=0.1)
    pyautogui.typewrite(input_word, interval=0.1)
    keyboard.press_and_release("enter")


# ------------------------------变量区------------------------------

nc_path = "E:\\NC\\UClient.exe"
# 查找图片存放的路径
target_image_path = "E:\\NC\\picture"
# 截图路径
origin_image_path = "E:\\NC\\picture\\origin"
# 用户名
username = ""
# 密码
password = ""
# ------------------------------变量区------------------------------


# ------------------------------主流程------------------------------

# 准备工作，选择目录，没有则创建
target_image_path = init_path(target_image_path)
origin_image_path = init_path(origin_image_path)

# 打开用友
process = subprocess.Popen(nc_path)

# 点击否
find_picture_click(origin_image_path, target_image_path, "not_origin.jpg", "not.jpg", "single")
time.sleep(1)

# 点击用友
find_picture_click(origin_image_path, target_image_path, "nc_origin.jpg", "nc.jpg", "double")
time.sleep(10)

# 输入用户名
find_picture_click_and_input(origin_image_path, target_image_path, "username_orgin.jpg", "username.jpg", username)
time.sleep(1)

# 输入密码，登录
find_picture_click_and_input(origin_image_path, target_image_path, "password_orgin.jpg", "password.jpg", password)
time.sleep(5)

process.wait()

# 点击会计
