"""
功能：
    查找template_path的图片在image_path这种图片下的xy坐标，并移动鼠标点击 
参数示例：
     image_path = "E:\\NC\\picture\\img.jpg"
     template_path = "E:\\NC\\picture\\template.jpg"
安装依赖：
      E:\\UERPA\\UERPA\\python-3.7.3-embed-amd64\\python.exe  -m pip install pyautogui
"""
import cv2
import pyautogui
from pynput import mouse
from PIL import ImageGrab

"""
查找image_path图中，template_path目标图像的坐标
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
计算出鼠标要点击的x/y坐标
"""


def get_click_point(top_left, bottom_right):
    # 鼠标左键点击的坐标
    x = (top_left[0] + bottom_right[0]) / 2
    y = (top_left[1] + bottom_right[1]) / 2
    return x, y


# 根据xy坐标点击
def on_click(x, y):
    # 单击鼠标左键
    # 参数1：移动x坐标到指定位置
    # 参数2：移动y坐标到指定位置
    # 参数3：点击次数
    # 无参数：在当前坐标单击
    pyautogui.click(x, y)


"""
# 查找图片并点击
# orginal_path 存放源目标的图片
# target_path 目标图片
# target_file_name 文件名
# action : single: 单击 , double: 双击, right: 右键
"""


def find_picture_click(orginal_path, target_path, target_file_name, action, extra_x=0, extra_y=0):
    # 找到最后一个圆点的索引位置
    last_dot_index = target_file_name.rfind(".")
    if last_dot_index != -1:
        file_name = target_file_name[:last_dot_index]  # 获取文件名部分
        file_extension = target_file_name[last_dot_index + 1:]  # 获取文件名后缀部分
    else:
        print("无法解析文件名和文件名后缀")
    # 截图的保存路径的绝对路径名
    orginal_path_name = orginal_path + file_name + "_origin." + file_extension
    # 目标图片的保存路径的绝对路径名
    target_path_name = target_path + file_name + "." + file_extension

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


# -------------------------------使用------------------------------- #
origin_image_path = "E:\\NC\\picture\\origin\\"
target_image_path = "E:\\NC\\picture\\"
find_picture_click(origin_image_path, target_image_path, "ask.jpg", "double",+50)

# PS 附录
'''
鼠标控制
# 获取鼠标当前坐标位置
pyautogui.position()

# 获取屏幕分辨率
pyautogui.size()

# 绝对坐标移动鼠标
# 参数1：移动x坐标到指定位置
# 参数2：移动y坐标到指定位置
# 参数3：移动速度，单位：秒
pyautogui.moveTo(参数1, 参数2, 参数3)

# 相对坐标移动鼠标
# 参数1：x坐标偏移量，正数：向右移动，负数：向左移动
# 参数2：y坐标偏移量，正数：向下移动，负数：向上移动
# 参数3：移动速度，单位：秒
pyautogui.moveRel(20, 50, 3)

# 按下按键
# 参数1：移动x坐标到指定位置
# 参数2：移动y坐标到指定位置
# 无参数：在当前坐标按下
pyautogui.mouseDown(300, 400)

# 释放按键
# 参数1：移动x坐标到指定位置
# 参数2：移动y坐标到指定位置
# 无参数：在当前坐标释放
pyautogui.mouseUp(400, 500)      

# 单击鼠标左键
# 参数1：移动x坐标到指定位置
# 参数2：移动y坐标到指定位置
# 参数3：点击次数
# 无参数：在当前坐标单击
pyautogui.click(参数1, 参数2, 参数3)

# 双击鼠标左键
# 参数1：移动x坐标到指定位置
# 参数2：移动y坐标到指定位置
# 无参数：在当前坐标双击
pyautogui.doubleClick(参数1, 参数2)

# 点击鼠标右键
# 参数1：移动x坐标到指定位置
# 参数2：移动y坐标到指定位置
# 无参数：在当前坐标点击
pyautogui.rightClick(参数1, 参数2)

#点击鼠标中键
# 参数1：移动x坐标到指定位置
# 参数2：移动y坐标到指定位置
# 参数3：移动速度，单位：秒
# duration=2（指定需要赋值的参数）
pyautogui.middleClick(100, 500, duration=2)

# 绝对坐标拖动鼠标（默认按住鼠标左键开始移动）
# 参数1：移动x坐标到指定位置
# 参数2：移动y坐标到指定位置
# 参数3：移动速度，单位：秒
# 参数4：left：按住鼠标左键 right：按住鼠标右键
pyautogui.dragTo(300, 400, 5)

# 相对坐标拖动鼠标（默认按住鼠标左键开始移动）
# 参数1：x坐标偏移量
# 参数2：y坐标偏移量
# 参数3：移动速度，单位：秒
# 参数4：left：按住鼠标左键 right：按住鼠标右键
pyautogui.dragRel(参数1, 参数2, 参数3, 参数4)

# 鼠标滚轮上下滚动
# 参数为上下滚动量，正数为向上滚动，负数为向下滚动。
pyautogui.scroll(300)


三、键盘控制
# 按下按键不放
pyautogui.keyDown("w")

# 释放按键
pyautogui.keyUp("w")

# 按下按键后立即释放
# 等同于调用keyDown("w")和keyUp("w")功能
pyautogui.press("w")

# 输入内容
# 参数1：需要输入的内容 可以为字符串("Python") 也可以为数组(["P","y","t","h","o","n"])
# 参数2：输入字符间隔时间
pyautogui.typewrite(参数1, 参数2)

# 组合键
# 如 全选("ctrl", "a") 复制("ctrl", "c") 粘贴("ctrl", "v")
pyautogui.hotkey("ctrl", "a")              # 全选
pyautogui.hotkey('ctrl', 'shift', 'esc')  # 调出任务管理器

# 输入汉字解决方案
pyperclip.copy("你好，可以通过我输入中文哦，哈哈")  # 复制需要输入的内容
pyautogui.hotkey("ctrl", "v")   # 粘贴内容
'''
