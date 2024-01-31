# import tkinter as tk
# from tkinter import messagebox
# 
# from tkinter.filedialog import askopenfilename, askdirectory
# 
# root = tk.Tk()
# root.withdraw()
# # 
# # # 弹出警告提示框
# result = messagebox.askquestion("警告", "这是一个警告提示框")
# if result == 'yes':
#     print("用户点击了确认按钮")
# else:
#     print("用户点击了取消按钮")
import pygame
from pygame.locals import *


def process_data_confirm(ttf_path):
    # 初始化 Pygame
    pygame.init()

    # 设置显示窗口大小
    screen = pygame.display.set_mode((300, 120))
    pygame.display.set_caption('协同确定窗口')
    # 定义颜色
    BLACK = (0, 0, 0)
    WHITE = (255, 255, 255)
    GRAY = (150, 150, 150)

    # 定义字体
    # 加载中文字体
    font = pygame.font.Font(ttf_path, 24)
    # 定义文本
    text_content = "请检查用友的数据是否正确"
    # 创建文本对象
    text = font.render(text_content, True, BLACK)

    # 定义按钮
    button_rect = pygame.Rect(50, 50, 100, 50)
    button_text = font.render("确定", True, BLACK)

    # 取消按钮
    cancel_button_rect = pygame.Rect(180, 50, 100, 50)
    cancel_button_text = font.render("取消", True, BLACK)

    # 游戏循环
    running = True
    while running:
        for event in pygame.event.get():
            if event.type == QUIT:
                running = False
                return False
            elif event.type == MOUSEBUTTONDOWN:
                # 检查是否点击了按钮
                if button_rect.collidepoint(event.pos):
                    # 点击了按钮，继续执行
                    print("用户点击了确定按钮,确定插入数据没问题")
                    running = False
                    return True
                elif cancel_button_rect.collidepoint(event.pos):
                    # 点击了取消按钮，继续执行
                    print("用户点击了取消按钮,插入数据有问题")
                    running = False
                    return False

        # 绘制背景
        screen.fill(WHITE)

        # 绘制文本
        text_rect = text.get_rect()
        screen.blit(text, text_rect)

        # 绘制按钮
        pygame.draw.rect(screen, GRAY, button_rect)
        pygame.draw.rect(screen, BLACK, button_rect, 1)
        button_text_rect = button_text.get_rect(center=button_rect.center)
        screen.blit(button_text, button_text_rect)

        # 绘制取消按钮
        pygame.draw.rect(screen, GRAY, cancel_button_rect)
        pygame.draw.rect(screen, BLACK, cancel_button_rect, 1)
        cancel_button_text_rect = cancel_button_text.get_rect(center=cancel_button_rect.center)
        screen.blit(cancel_button_text, cancel_button_text_rect)

        # 更新屏幕
        pygame.display.update()

    pygame.quit()
