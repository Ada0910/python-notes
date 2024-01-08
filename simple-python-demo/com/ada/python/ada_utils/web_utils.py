# 启动Chrome浏览器
import os
import platform

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

"""
获取谷歌浏览器驱动的路径

Args:
    filename：文件名，全路径
Returns:
    返回：谷歌浏览器驱动的路径

Raises:
    异常情况的说明（如果有的话）

"""


def get_driver_location():
    os_type = platform.system()
    root_dir = os.path.dirname(os.path.abspath(__file__))
    drivers_dir = os.path.join(root_dir, 'drivers')
    print("当前的环境的操作系统是：【" + os_type + "】，根目录是：【" + root_dir + "】，谷歌浏览器驱动的路径是：【" + drivers_dir + "】")
    if os_type == 'Darwin':
        return os.path.join(drivers_dir, 'chromedriver_mac64')
    elif os_type == 'Windows':
        return os.path.join(drivers_dir, 'chromedriver.exe')
    elif os_type == 'Linux':
        return os.path.join(drivers_dir, 'chromedriver_linux64')
    else:
        return None


"""

创建 WebDriver 实例并连接到已经打开的浏览器

方法一：
需要先启动一个浏览器，并将其启动参数 `remote-debugging-port` 设置为可用的端口号
例如 Chrome 浏览器可以在命令行中使用 `--remote-debugging-port=9222` 参数启动
确保已经打开的浏览器支持远程调试，并且与指定的端口号匹配
"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe --remote-debugging-port=9222"


方法二：
快捷方式添加 --remote-debugging-port=9222 --user-data-dir="E:\\PythonCode\\chrome-dir"
"""


def get_exist_chrome():
    # 创建 WebDriver 实例并连接到已经打开的浏览器
    # 需要先启动一个浏览器，并将其启动参数 `remote-debugging-port` 设置为可用的端口号
    # 例如 Chrome 浏览器可以在命令行中使用 `--remote-debugging-port=9222` 参数启动
    # 确保已经打开的浏览器支持远程调试，并且与指定的端口号匹配
    # C:\Program Files\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222
    # service = Service(executable_path=MajorChromeDriverManager().install())
    # 快捷方式添加 --remote-debugging-port=9222 --user-data-dir="E:\PythonCode\chrome-dir"
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    return webdriver.Chrome(options=chrome_options)


"""
创建 WebDriver 实例并打开一个新的浏览器实例
"""


def get_new_chrome():
    # 检测类型
    driver_location = get_driver_location()
    print("谷歌浏览器的路径：{}".format(driver_location))
    if driver_location is None:
        print('不支持的系统类型！')
        exit(-1)

    opt = Options()
    opt.binary_location = driver_location
    return webdriver.Chrome(options=opt)