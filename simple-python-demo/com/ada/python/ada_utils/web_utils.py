# 启动Chrome浏览器
import os
import platform

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
from selenium.webdriver.chrome.service import Service

"""
获取谷歌浏览器驱动的路径

Args:
    filename：文件名，全路径
Returns:
    返回：谷歌浏览器驱动的路径

Raises:
    异常情况的说明（如果有的话）

"""


def get_chrome_driver_location():
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

参数实例
# chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# remote_debugging_port = 9222
# user_data_dir = r"E:\PythonCode\chrome-dir"
"""


def get_exist_chrome_driver(chrome_path, remote_debugging_port, user_data_dir):
    # 创建 WebDriver 实例并连接到已经打开的浏览器

    # 构建命令行命令
    command = f'"{chrome_path}" --remote-debugging-port={remote_debugging_port} --user-data-dir="{user_data_dir}"'
    # 执行命令
    subprocess.Popen(command, shell=True)
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    return webdriver.Chrome(options=chrome_options)


"""
创建 WebDriver 实例并打开一个新的浏览器实例
"""


def get_new_chrome_driver():
    # 检测类型
    driver_location = get_chrome_driver_location()
    print("谷歌浏览器的路径：{}".format(driver_location))
    if driver_location is None:
        print('不支持的系统类型！')
        exit(-1)

    opt = Options()
    opt.binary_location = driver_location
    return webdriver.Chrome(options=opt)


"""
创建一个edge实例并打开
"""


def get_new_edge_driver(msedgedriver_path_name, download_path):
    opt = webdriver.EdgeOptions()
    # 设置自动下载文件选项
    opt.add_argument("--download.prompt_for_download=false")  # 禁止弹出下载对话框
    opt.add_argument("--browser.download.folderList=2")  # 禁止弹出下载对话框
    opt.add_argument("--download.default_directory=" + download_path)  # 设置下载文件的保存路径
    prefs = {
        'download.default_directory': download_path}
    opt.add_experimental_option("prefs", prefs)
    # 创建 Edge 服务
    service = Service(msedgedriver_path_name)
    return webdriver.Edge(service=service, options=opt)
