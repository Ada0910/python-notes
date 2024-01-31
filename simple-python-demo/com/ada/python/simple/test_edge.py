import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
import os

# 设置Edge浏览器驱动的路径
# driver_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
from selenium.webdriver.edge.service import Service

driver_path = "D:\\PythonCode\\edge\\msedgedriver.exe"

# 创建一个EdgeOptions对象
options = Options()

# 设置Edge浏览器下载文件的默认路径
prefs = {
    "download.default_directory": "D:\\PythonCode\\edge",  # 设置为当前工作目录
    "download.prompt_for_download": False,  # 不询问下载路径
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False,
    "plugins.always_open_pdf_externally": True  # 如果下载的是PDF文件，始终在外部程序中打开
}

# 将设置添加到EdgeOptions对象
options.add_experimental_option("prefs", prefs)

s = Service(driver_path)
driver = webdriver.Edge(service=s)


# 打开JDK 8官网
driver.get('https://developer.microsoft.com/zh-cn/microsoft-edge/tools/webdriver/?form=MA13LH')

time.sleep(20)
# print(driver.page_source)
elements = driver.find_elements(By.CSS_SELECTOR, "div.block-web-driver__version-links > a")
print("elementt",elements[0].text)
elements[0].click()

# 点击下载按钮，进行文件下载操作
# ...
time.sleep(6666)
# 关闭浏览器
# driver.quit()
