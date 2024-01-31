"""
驱动下载地址：http://chromedriver.storage.googleapis.com/index.html
"""
import time

from selenium import webdriver
import os
import platform
import pytesseract
from PIL import Image
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pyautogui
import subprocess
import pychrome
from selenium.webdriver.chrome.service import Service
import requests

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


def get_exist_chrome(chrome_path, user_data_dir):
    remote_debugging_port = 9222
    # 构建命令行命令
    command = f'"{chrome_path}" --remote-debugging-port={remote_debugging_port} --user-data-dir="{user_data_dir}"'
    # 执行命令
    subprocess.Popen(command, shell=True)
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-default-apps")
    # prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': 'e:\\'}
    # chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    return webdriver.Chrome(options=chrome_options)


def get_chrome():
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': 'e:\\'}
    options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome(executable_path='C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
                              chrome_options=options)
    return driver


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
    opt = webdriver.ChromeOptions()

    # opt.add_argument('--disable-extensions')
    # opt.add_argument('--disable-infobars')
    # opt.add_argument('--disable-popup-blocking')
    # opt.add_argument('--disable-save-password-bubble')
    # opt.add_argument('--disable-translate')
    # opt.add_argument('--ignore-certificate-errors')
    # opt.add_argument("--safebrowsing-disable-download-protection")
    # opt.add_argument('--start-maximized')
    # opt.add_argument('--disable-web-security')
    # opt.add_argument("--headless")  # 无头模式，可选
    opt.add_argument("--disable-gpu")  # 禁用GPU加速，可选
    opt.add_argument("--no-sandbox")  # 非沙盒模式，可选
    opt.add_argument("--disable-popup-blocking")  # 禁用弹窗拦截，可选
    opt.add_argument("--disable-infobars")  # 禁用信息栏，可选
    prefs = {
        # 设置下载路径
        'download.default_directory': 'E:\\NC\\terminal\\template3\\',
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        "safebrowsing.disable_download_protection": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "download_restrictions": 0,
        "browser.download.folderList": 1,
        # 设置为 2 禁止弹出窗口
        "profile.default_content_settings.popups": 2
    }
    opt.add_experimental_option("prefs", prefs)
    opt.binary_location = driver_location
    return webdriver.Chrome(options=opt)


"""
打印出当前页面的元素
"""


def print_content(driver):
    # 获取页面内容
    page_content = driver.page_source
    print("当前的page_content是：", page_content)


"""
等待新窗口打开并切换到新窗口

Args:
    driver：谷歌浏览器实例
Returns:
    返回：无

Raises:
    异常情况的说明（如果有的话）
"""


def switch_to_new_window(driver):
    # 获取当前窗口句柄
    current_window_handle = driver.current_window_handle
    # 等待新窗口打开并切换到新窗口
    driver.implicitly_wait(10)
    all_window_handles = driver.window_handles
    new_window_handle = [handle for handle in all_window_handles if handle != current_window_handle][0]
    driver.switch_to.window(new_window_handle)


def get_new_edge():
    msedgedriver_path = "C:\\Users\\UMS-AUG\\Downloads\\msedgedriver.exe"
    opt = Options()
    opt = webdriver.EdgeOptions()
    download_path = "E:\\NC\\terminal\\template\\"
    # 设置自动下载文件选项
    opt.add_argument("--download.default_directory="+download_path)  # 设置下载文件的保存路径
    opt.add_argument("--download.prompt_for_download=false")  # 禁止弹出下载对话框
    opt.add_argument("--browser.download.folderList=2")  # 禁止弹出下载对话框
    # prefs = {
    #     # 设置下载路径
    #     'download.default_directory': 'E:\\NC\\terminal\\template3\\',
    #     # "download.prompt_for_download": False,  # 不询问下载路径
    #     # "download.directory_upgrade": True,
    #     # "safebrowsing.enabled": False
    #     # 设置为 2 禁止弹出窗口
    #     # # 置成 2 表示使用自定义下载路径；设置成 0 表示下载到桌面；设置成 1 表示下载到默认路径
    #     # "browser.download.folderList":2,
    #     # "browser.download.manager.showWhenStarting": False,
    #     # "profile.default_content_settings.popups": 2
    # }
    # 
    # opt.add_experimental_option("prefs", prefs)
    # 创建 Edge 服务
    service = Service(msedgedriver_path)
    driver = webdriver.Edge(service=service, options=opt)
    return driver


def process_get_edge_driver():
    user_data_dir = r"E:\PythonCode\edge-dir"
    remote_debugging_port = 9222
    # 设置 Edge 驱动程序路径
    edge_driver_path = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'  # 请替换为你的 Edge 驱动程序路径

    # 设置 Edge 浏览器下载路径
    download_path = "C:/Downloads"  # 请替换为你想要设置的下载路径
    dirver_path = "C:\\Users\\UMS-AUG\\Downloads\\msedgedriver.exe"

    option = webdriver.EdgeOptions()
    option.add_experimental_option("debuggerAddress", "127.0.0.1:9527")
    
    driver = webdriver.Edge(options=option)
    # 创建 Edge 配置选项
    options = Options()
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    options.add_experimental_option('prefs', prefs)
    # 创建 Edge 服务
    service = Service(dirver_path)

    # # 构建命令行命令
    # command = f'"{edge_driver_path}" --remote-debugging-port={remote_debugging_port} --user-data-dir="{user_data_dir}"'
    # # 执行命令
    # subprocess.Popen(command, shell=True)
    # 启动 Edge 浏览器
    driver = webdriver.Edge(service=service, options=options)
    return driver


def process_to_am(driver):
    # 最大化
    # driver.maximize_window()
    driver.get("https://sso.chinaums.com/selfService/welcome");
    print("【主流程】谷歌浏览器启动成功，当前的URL是：", driver.current_url)
    wait = WebDriverWait(driver, 300)
    content = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    # 定位到frame标签
    driver.switch_to.frame("mainFrame")
    print("准备定位资产管理系统的图标的元素")
    # 点击资产管理系统
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, "li.app-list-item > p")
        for element in elements:
            title = element.text.strip()
            if title == '天码服务平台':
                element.click()
                break
    except Exception:
        print("页面找不到资产管理系统的图标的元素！")
    else:
        print("正在跳转到资产管理系统.......")
        # 等待新窗口打开并切换到新窗口
        switch_to_new_window(driver)
        # 睡眠
        time.sleep(3)
        element = driver.find_element(By.CSS_SELECTOR, "div.cell > a")
        print(element.text)
        element.click()
        print(element.text)

        # export_link = driver.find_element(By.CSS_SELECTOR, "div.cell > a")
        # export_url = export_link.get_attribute("href")
        # print("导出的URL是：",export_url)

        # template_path_name = "E:\\NC\\terminal\\template\\1.xlsx"
        # # response = requests.get(export_url)
        # # with open(template_path_name, 'wb') as f:
        # #     f.write(response.content)
        # file_content = driver.page_source.encode('utf-8')
        # print("file_content", file_content)
        # with open(template_path_name, 'wb') as f:
        #     f.write(file_content)

        # alert = driver.switch_to.alert
        # # 获取弹窗文本并打印
        # print(alert.text)
        # alert.accept()

        # # 按下ctrl+J 查看下载内容
        # # 执行 JavaScript 代码模拟按下 Ctrl+J 快捷键
        # # driver.get("chrome://settings/")
        # driver.execute_script("window.open('chrome://downloads/', '_blank');")
        # driver.get("chrome://settings/")
        # pyautogui.click(0, 0)
        # pyautogui.hotkey("ctrl", "j")
        # time.sleep(1)
        # # 获取打开的多个窗口句柄
        # windows = driver.window_handles
        # # 切换到当前最新打开的窗口
        # driver.switch_to.window(windows[-1])
        # element = driver.find_elements(By.TAG_NAME, "cr-button")
        # print(element[0].text)
        # element[1].click()
        # 
        # # 关闭当前标签页
        # driver.close()
        # # 切换到当前最新打开的窗口
        # driver.switch_to.window(windows[-1])
        # time.sleep(1)
        # time.sleep(66667)


def print_content(driver):
    # 获取页面内容
    page_content = driver.page_source
    print("当前的page_content是：", page_content)


def open_new_tab(driver):
    driver.execute_script("""
        var link = document.createElement('a');
        document.body.appendChild(link);
        link.href = "http://10.11.32.226:8080/am/export/updateLostDetailBatch.xlsx";
        link.click();
        document.body.removeChild(link);
    """)
    # 下载内容
    # driver.execute_cdp_cmd("Page.navigate", {"url": "chrome://downloads/"})


def download_captcha_image(driver):
    import base64
    element = driver.find_element(By.CSS_SELECTOR, "div#img-code > img")
    save_path = 'E:\\a.jpeg'
    # 获取img标签的src内容
    src = element.get_attribute('src')
    base64_data = src.split(',')[1]
    binary_data = base64.b64decode(base64_data)
    # 将二进制数据保存为jpg文件
    with open('output.jpg', 'wb') as f:
        f.write(binary_data)


if __name__ == "__main__":
    # driver = get_chrome()
    # driver = get_new_chrome()
    # driver = process_get_edge_driver()
    driver = get_new_edge()

    # # 谷歌浏览路径
    # chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    # 
    # # 谷歌用户目录
    # user_data_dir = r"E:\PythonCode\chrome-dir"
    # driver = get_exist_chrome(chrome_path, user_data_dir)
    # driver.maximize_window()
    process_to_am(driver)
    # # chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    # # user_data_dir = r"E:\PythonCode\chrome-dir"
    # # driver = get_exist_chrome(chrome_path=chrome_path, user_data_dir=user_data_dir)
    time.sleep(1233)
    # # # https://pay.gdrcu.com/userPlatform/#/
