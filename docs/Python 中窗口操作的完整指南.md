# Python 中窗口操作的完整指南

## 1. 使用 `pygetwindow` 库获取窗口列表

`pygetwindow` 库提供了获取窗口列表和操作窗口的功能。

```
import pygetwindow as gw

# 获取当前打开的所有窗口
all_windows = gw.getWindowsWithTitle('')
for window in all_windows:
    print(window)
```

## 2. 使用 `pygetwindow` 将窗口置顶

可以使用 `pygetwindow` 将窗口置顶显示。

```
import pygetwindow as gw

# 获取指定窗口并置顶
window = gw.getWindowsWithTitle('Your Window Title')[0]
window.setAlwaysOnTop(True)
```

## 3. 使用 `pygetwindow` 最大化和最小化窗口

使用 `pygetwindow` 库可以轻松地将窗口最大化或最小化。

```
import pygetwindow as gw

# 获取指定窗口并最大化
window = gw.getWindowsWithTitle('Your Window Title')[0]
window.maximize()

# 最小化窗口
window.minimize()
```

## 4. 使用 `pygetwindow` 移动窗口到指定位置

可以将窗口移动到屏幕的指定位置。

```
import pygetwindow as gw

# 获取指定窗口并移动到指定位置
window = gw.getWindowsWithTitle('Your Window Title')[0]
window.moveTo(100, 100)  # 移动到 x=100, y=100 的位置
```

## 5. 使用 `pygetwindow` 获取窗口的大小和位置

`pygetwindow` 库允许获取窗口的大小和位置。

```
import pygetwindow as gw

# 获取指定窗口的大小和位置
window = gw.getWindowsWithTitle('Your Window Title')[0]
print(window.size)   # 获取窗口大小
print(window.left, window.top)  # 获取窗口左上角位置
```

## 6. 使用 `pygetwindow` 激活并关闭窗口

可以使用 `pygetwindow` 激活窗口并将其关闭。

```
import pygetwindow as gw

# 获取指定窗口并激活
window = gw.getWindowsWithTitle('Your Window Title')[0]
window.activate()

# 关闭窗口
window.close()
```

## 7. 使用 `pyautogui` 获取屏幕分辨率

`pyautogui` 库可用于获取屏幕的分辨率。

```
import pyautogui

# 获取屏幕分辨率
screen_width, screen_height = pyautogui.size()
print(f"屏幕分辨率: {screen_width}x{screen_height}")
```

## 8. 使用 `pyautogui` 获取鼠标当前位置

可以利用 `pyautogui` 获取鼠标当前的位置。

```
import pyautogui

# 获取鼠标当前位置
current_x, current_y = pyautogui.position()
print(f"鼠标位置: x={current_x}, y={current_y}")
```

## 9. 使用 `pyautogui` 移动鼠标和点击

`pyautogui` 可以模拟鼠标移动和点击。

```
import pyautogui

# 移动鼠标到指定位置
pyautogui.moveTo(100, 100, duration=1)  # 移动到 x=100, y=100 的位置，持续 1 秒

# 模拟鼠标点击
pyautogui.click()
```

## 10. 使用 `pyautogui` 模拟键盘输入

`pyautogui` 还可以模拟键盘输入。

```
import pyautogui

# 输入字符串
pyautogui.typewrite("Hello, World!", interval=0.1)  # 每个字符间隔 0.1 秒
```

## 11. 使用 `win32gui` 获取窗口句柄

`win32gui` 库可用于获取窗口的句柄。

```
import win32gui

# 获取窗口句柄
hwnd = win32gui.FindWindow(None, 'Your Window Title')
print(hwnd)
```

## 12. 使用 `win32gui` 获取窗口大小和位置

`win32gui` 还可用于获取窗口的大小和位置。

```
import win32gui

# 获取窗口大小和位置
rect = win32gui.GetWindowRect(hwnd)
print(f"窗口位置: {rect}")
```

## 13. 使用 `win32gui` 将窗口置顶

`win32gui` 可以帮助你将窗口置顶。

```
import win32gui
import win32con

# 将窗口置顶
win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
```

## 14. 使用 `win32gui` 最大化和最小化窗口

利用 `win32gui` 可以将窗口最大化或最小化。

```
import win32gui
import win32con

# 最大化窗口
win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

# 最小化窗口
win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
```

## 15. 使用 `win32gui` 移动窗口到指定位置

`win32gui` 可以将窗口移动到屏幕的指定位置。

```
import win32gui

# 移动窗口到指定位置
win32gui.SetWindowPos(hwnd, None, 100, 100, 0, 0, win32con.SWP_NOSIZE)
```

## 16. 使用 `win32api` 获取缩放比例

`win32api` 可以帮助你获取窗口的缩放比例。

```
import win32api

# 获取缩放比例
scaling_factor = win32api.GetScaleFactorForDevice(0)  # 0 表示主显示器
print(f"缩放比例: {scaling_factor}")
```

以上示例展示了如何使用不同的 Python 库来操纵窗口、获取窗口信息、控制鼠标和键盘，并获取屏幕信息。这些功能可帮助你实现各种窗口操作和自动化任务。

## 总结

本指南深入探讨了如何利用 Python 中的各种库来操纵窗口和执行窗口操作。通过 `pygetwindow` 库，分享了如何获取窗口列表、将窗口置顶、最大化、最小化以及移动到指定位置。`pyautogui` 库能够获取屏幕分辨率、鼠标位置，并模拟鼠标移动、点击和键盘输入。使用 `win32gui` 和 `win32api` 库，了解了如何获取窗口句柄、设置窗口大小、位置、置顶，最大化、最小化，并获取窗口的缩放比例。

这些示例提供了全面的指南，展示了如何利用 Python 中的多个库执行各种窗口操作，包括自动化任务、获取窗口信息和控制窗口外观。这些技巧和工具可帮助开发者在实现自动化脚本、进行窗口级别操作或执行定制化任务时更加灵活和高效。通过掌握这些方法，可以更好地理解和利用 Python 中丰富的窗口操控功能

**PS：转载https://www.sibida.net/detail/1990**
