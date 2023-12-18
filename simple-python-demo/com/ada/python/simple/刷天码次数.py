import keyboard
import time
import  pyautogui

# 提高采纳率
for num in range(0, 50):
    time.sleep(1)
    pyautogui.typewrite("System", interval=0.1)
    keyboard.press_and_release("space")
    time.sleep(1.5)
    # 按tab键
    keyboard.press_and_release("Tab")
    time.sleep(1)
    # 按回车
    keyboard.press_and_release("enter")


# 降低采纳率
for num in range(0, 100):
    time.sleep(1)
    pyautogui.typewrite("print", interval=0.1)
    keyboard.press_and_release("space")
    time.sleep(1.5)
    # 按回车
    keyboard.press_and_release("enter")