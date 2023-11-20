"""
Python—-pywin32如何获取窗口句柄
"""
import sys
import win32gui
import win32con


# .获取句柄的标题
def get_title(hwnd):
    title = win32gui.GetWindowText(hwnd)
    print('窗口标题:%s' % (title))
    return title


# 获取窗口类名
def get_clasname(hwnd):
    clasname = win32gui.GetClassName(hwnd)
    print('窗口类名:%s' % (clasname))
    return clasname


# 获取所有窗口的句柄
def get_all_windows():
    hWnd_list = []
    win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWnd_list)
    print(hWnd_list)
    return hWnd_list


# 获取窗口的子窗口句柄
def get_son_windows(parent):
    hWnd_child_list = []
    win32gui.EnumChildWindows(parent, lambda hWnd, param: param.append(hWnd), hWnd_child_list)
    print(hWnd_child_list)
    for i in hWnd_child_list:
        get_clasname(i)
        # get_son_windows(i)
    return hWnd_child_list


# 窗口置顶
def set_top(hwnd):
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER | win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE)


# 窗口取消置顶
def set_down(hwnd):
    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                          win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)


# 根据窗口名称获取句柄
def get_hwnd_from_name(name):
    hWnd_list = get_all_windows()
    for hwd in hWnd_list:
        title = get_title(hwd)
        if title == name:
            return hwd


# 窗口显示
def xianshi(name):
    hwd = get_hwnd_from_name(name)
    win32gui.ShowWindow(hwd, win32con.SW_SHOW)


# 窗口隐藏
def yingcang(name):
    hwd = get_hwnd_from_name(name)
    win32gui.ShowWindow(hwd, win32con.SW_HIDE)

# 获取右下角托盘的任务句柄
def get_tuopan_hwd():
    handle = win32gui.FindWindow("Shell_TrayWnd", None)
    hWnd_child_list = get_son_windows(handle)[1:]
    tuopan_hwd_list = []
    flag = False
    for i in hWnd_child_list:
        if get_clasname(i) == 'TrayNotifyWnd':
            flag = True
        if flag:
            tuopan_hwd_list.append(i)
    return tuopan_hwd_list


# 隐藏托盘
def yingcang(name=''):
    tuopan_hwd_list = get_tuopan_hwd()
    if name == '':
        for i in tuopan_hwd_list[:7]:  # [：7]因为要保留一些基本的内容，也可以全部隐藏
            win32gui.ShowWindow(i, win32con.SW_HIDE)
    else:
        win32gui.ShowWindow(name, win32con.SW_HIDE)


# 显示托盘
def xianshi(name = ''):
    tuopan_hwd_list = get_tuopan_hwd()
    if name == '':
        for i in tuopan_hwd_list:
            win32gui.ShowWindow(i, win32con.SW_SHOW)
    else:
        win32gui.ShowWindow(name, win32con.SW_SHOW)

