import keyboard
import time

for num in range(0, 50):
    time.sleep(1)
    # 按tab键
    keyboard.press_and_release("Tab")
    time.sleep(1)
    # 按回车
    keyboard.press_and_release("enter")
