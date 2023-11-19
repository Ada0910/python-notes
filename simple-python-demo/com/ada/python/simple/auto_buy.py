import os
import platform
import time
import datetime
from selenium import webdriver


# 启动Chrome浏览器
def get_driver():
    os_type = platform.system()
    root_dir = os.path.dirname(os.path.abspath(__file__))
    drivers_dir = os.path.join(root_dir, '../../../../../docs/drivers')
    print("当前的环境的操作系统是：【" + os_type + "】，根目录是：【" + root_dir + "】，谷歌浏览器驱动的路径是：【" + drivers_dir + "】")
    if os_type == 'Darwin':
        return os.path.join(drivers_dir, 'chromedriver_mac64')
    elif os_type == 'Windows':
        return os.path.join(drivers_dir, 'chromedriver.exe')
    elif os_type == 'Linux':
        return os.path.join(drivers_dir, 'chromedriver_linux64')
    else:
        return None


# 打开链接
def open_website(chrome, url):
    chrome.get(url)
    print("打开的链接是：【"+url+"】")
    chrome.implicitly_wait(10)


'''
下单函数
store = 1:淘宝
store = 2:天猫
store = 3:天猫超市
'''


def buy(chrome, store, buy_time):
    # 淘宝
    if store == '1':
        # "立即购买"的css_selector
        btn_buy = '#J_juValid > div.tb-btn-buy > a'
        # "立即下单"的css_selector
        btn_order = '#submitOrderPC_1 > div.wrapper > a'
        print("抢购的平台是：【淘宝】")
    # 天猫
    elif store == '2':
        btn_buy = '#J_LinkBuy'
        btn_order = '#submitOrderPC_1 > div > a'
        print("抢购的平台是：【天猫】")
    # 天猫超市
    elif store == '3':
        btn_buy = '#J_Go'
        btn_order = '#submitOrderPC_1 > div > a.go-btn'
        print("抢购的平台是：【天猫超市】")

    print("开始抢购，时间是【"+buy_time+"】")

    # for i in range(10000): # 刷新次数
    #     chrome.refresh()  # 刷新网页
    #     time.sleep(0.1) # 五秒一次
    while True:
        # 现在时间大于预设时间则开售抢购
        if datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f') > buy_time:
            # 刷新网页
            chrome.refresh()
            try:
            # 找到"立即购买" 点击
                if chrome.find_element_by_css_selector(btn_buy):
                    chrome.find_element_by_css_selector(btn_buy).click()
                    break
                # time.sleep(0.1)
            except:
                print('----------点击【立即购买】有异常----------')

    while True:
        try:
            # 找到"立即下单" 点击
            if chrome.find_element_by_css_selector(btn_order):
                chrome.find_element_by_css_selector(btn_order).click()
                # 下单成功，跳转至支付页面
                print('=============================【购买成功】=============================')
                break
        except:
            print('----------点击【立即下单】有异常----------')


if __name__ == "__main__":
    # 检测类型
    driver_location = get_driver()
    if driver_location is None:
        print('不支持的系统类型！')
        exit(-1)
    chrome = webdriver.Chrome(driver_location)
    chrome.maximize_window()
    # 输入
    # url = input('请输入链接\n - 淘宝/天猫：输入商品链接；\n - 天猫超市：输入购物车链接并勾选预购商品;\n')
    # store = input('请输入商城序号:\n1. 淘宝\n2. 天猫\n3. 天猫超市\n')
    # buy_time = input('请输入抢购时间:\neg. 2022-02-19  19:40:00\n')

    url = 'https://detail.tmall.com/item.htm?spm=a1z10.1-b-s.w5003-24221126476.1.2b1f3e6d3UoPMn&id=606733853474&scene' \
          '=taobao_shop&skuId=4988581327279 '
    store = '2'
    buy_time = '2022-02-20 18:59:59'
    print('请手动登录（务必在抢购时间前完成）')
    # 先打开网址
    open_website(chrome, url)
    # 下单抢购
    buy(chrome, store, buy_time)
