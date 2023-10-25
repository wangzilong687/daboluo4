import time
import pyautogui
import win32gui
import os
import sys
import io
import cv2
import numpy as np
from PIL import ImageGrab
import random

from PIL import Image
import pytesseract

from aip import AipOcr

import configparser

# pyinstaller -D --add-data "config.ini;." --add-data "utils/;utils/"  daboluo.py

# 如果程序被打包，使用打包的路径
if getattr(sys, 'frozen', False):
    config_path = os.path.join(sys._MEIPASS, 'config.ini')
# 否则，使用当前路径
else:
    config_path = 'config.ini'

config = configparser.ConfigParser()
config.read(config_path, encoding='utf-8')

file_path = config.get('DEFAULT', 'file_path')
start_x = int(config.get('DEFAULT', 'start_x'))
start_y = int(config.get('DEFAULT', 'start_y'))
x_count = int(config.get('DEFAULT', 'x_count'))
y_count = int(config.get('DEFAULT', 'y_count'))
x_grow = int(config.get('DEFAULT', 'x_grow'))
y_grow = int(config.get('DEFAULT', 'y_grow'))
box_num = int(config.get('DEFAULT', 'box_num'))
box_x = int(config.get('DEFAULT', 'box_x'))
box_y = int(config.get('DEFAULT', 'box_y'))
box_grow = int(config.get('DEFAULT', 'box_grow'))

kuzi = config.get('Keywords', 'kuzi').split(',')
xiezi = config.get('Keywords', 'xiezi').split(',')
toukui = config.get('Keywords', 'toukui').split(',')
xiongjia = config.get('Keywords', 'xiongjia').split(',')
shoutao = config.get('Keywords', 'shoutao').split(',')
jiezhi = config.get('Keywords', 'jiezhi').split(',')
hufu = config.get('Keywords', 'hufu').split(',')
shuangshouchui = config.get('Keywords', 'shuangshouchui').split(',')
chui = config.get('Keywords', 'chui').split(',')
fu = config.get('Keywords', 'fu').split(',')
shuangshoufu = config.get('Keywords', 'shuangshoufu').split(',')
jian = config.get('Keywords', 'jian').split(',')
shuangshoujian = config.get('Keywords', 'shuangshoujian').split(',')
changbing = config.get('Keywords', 'changbing').split(',')

# 死灵
liandao = config.get('Keywords', 'liandao').split(',')
shuangshouliandao = config.get('Keywords', 'shuangshouliandao').split(',')
junengqi = config.get('Keywords', 'junengqi').split(',')
dun = config.get('Keywords', 'dun').split(',')
bishou = config.get('Keywords', 'bishou').split(',')


def get_window(name):
   # GetDesktopWindow 获得代表整个屏幕的一个窗口（桌面窗口）句柄
    hd = win32gui.GetDesktopWindow()

    # 获取所有子窗口
    hwndChildList = []

    win32gui.EnumChildWindows(
        hd, lambda hwnd, param: param.append(hwnd), hwndChildList)
    for hwnd in hwndChildList:
        if (win32gui.GetWindowText(hwnd) == name):
            t_window = win32gui.FindWindow(
                0, win32gui.GetWindowText(hwnd))  # 寻找窗口
            if t_window:
                win32gui.SetForegroundWindow(t_window)  # 前置窗口


def save_image(img, dir_name, file_name):
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
    full_path = os.path.join(dir_name, file_name)
    img.save(full_path)


def find_text_position(template_path):
    # 截取整个屏幕
    img = ImageGrab.grab()

    # 将PIL.Image.Image对象转换为numpy.array对象
    screenshot = np.array(img)

    # 将RGB图像转换为BGR图像，因为OpenCV使用BGR格式
    screenshot = screenshot[:, :, ::-1]

    # 读取模板图像
    template = cv2.imread(template_path)

    # 使用OpenCV的matchTemplate函数进行模板匹配
    result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)

    # 获取匹配结果的最大值和最大值的位置
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # 如果匹配度大于80%
    if max_val > 0.7:
        # 直接返回模板图像在窗口中的坐标
        return max_loc
    else:
        return None


def baiduai(image):
    # 定义常量
    APP_ID = '41660875'
    API_KEY = 'BualnTKFOBm4a3hmpuX8oSx0'
    SECRET_KEY = 'P2B33hbi7Frl6c9aGHynyATBM4dLEGFQ'

    the_time = None
    the_name = None
    level = 'normal'

    keywords = None
    count = 0

    # 初始化AipOcr对象
    aipOcr = AipOcr(APP_ID, API_KEY, SECRET_KEY)

    # 读取图片
    def get_image_content(image):
        img_byte_arr = io.BytesIO()
        image.save(img_byte_arr, format='PNG')
        img_byte_arr = img_byte_arr.getvalue()
        return img_byte_arr

    # 定义参数变量
    options = {
        'detect_direction': 'true',
        'language_type': 'CHN_ENG',
    }

    # 调用通用文字识别接口
    result = aipOcr.basicGeneral(get_image_content(image), options)
    list = result['words_result']

    for item in list:
        if '先祖稀有裤子' in item['words']:
            the_name = '裤子'  # 生命上限
            keywords = kuzi
            break
        elif ('先祖稀有鞋子' in item['words']):
            the_name = '鞋子'
            keywords = xiezi
            break
        elif ('先祖稀有头盔' in item['words']):
            the_name = '头盔'  # 冷却 生命 护甲 怒气
            keywords = toukui
            break
        elif ('先祖稀有胸甲' in item['words']):
            the_name = '胸甲'  # 生命上限 总护甲 近距敌人 被强固时 压制伤害
            keywords = xiongjia
            break
        elif ('先祖稀有手套' in item['words']):
            the_name = '手套'  # 攻击速度 暴击几率 压制伤害 力量
            keywords = shoutao
            break
        elif ('先祖稀有戒指' in item['words']):
            the_name = '戒指'
            keywords = jiezhi
            break
        elif ('先祖稀有护符' in item['words']):
            the_name = '护符'
            keywords = hufu
            break
        elif ('先祖稀有双手锤' in item['words']):
            the_name = '双手锤'
            keywords = shuangshouchui
            break
        elif ('先祖稀有锤' in item['words']):
            the_name = '单手锤'
            keywords = chui
            break
        elif ('先祖稀有斧' in item['words']):
            the_name = '单手斧头'
            keywords = fu
            break
        elif ('先祖稀有双手斧' in item['words']):
            the_name = '双手斧'
            keywords = shuangshoufu
            break
        elif ('先祖稀有剑' in item['words']):
            the_name = '单手剑'
            keywords = jian
            break
        elif ('先祖稀有双手剑' in item['words']):
            the_name = '双手剑'
            keywords = shuangshoujian
            break
        elif ('先祖稀有长柄武器' in item['words']):
            the_name = '长柄'
            keywords = changbing
            break
        elif ('先祖稀有镰刀' in item['words']):
            the_name = '镰刀'
            keywords = liandao
            break
        elif ('先祖稀有双手镰刀' in item['words']):
            the_name = '双手镰刀'
            keywords = shuangshouliandao
            break
        elif ('先祖稀有聚能器' in item['words']):
            the_name = '聚能器'
            keywords = junengqi
            break
        elif ('先祖稀有盾' in item['words']):
            the_name = '盾'
            keywords = dun
            break
        elif ('先祖稀有匕首' in item['words']):
            the_name = '匕首'
            keywords = bishou
            break

    if keywords is not None:
        for keyword in keywords:
            for item in list:
                print(keyword,item['words'])
                if keyword in item['words']:
                    count += 1
                    break

    if count >= 3:
        level = '3中'
    elif count >= 2:
        level = '2中'
    elif count >= 1:
        level = '1中'
    else:
        level = '垃圾'

    print(count, 'count')
    return (the_name, level)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def _run_():
    # 设置鼠标移动到的位置，这里只是示例，你需要根据实际情况设置
    c_start_x = start_x
    c_start_y = start_y
    c_x_count = x_count
    c_y_count = y_count
    c_box_x = box_x
    c_box_y = box_y
    for b in range(box_num):
        pyautogui.moveTo(c_box_x, c_box_y, 0.3)
        pyautogui.click(c_box_x, c_box_y)
        c_start_y = start_y

        for y in range(c_y_count):
            for x in range(c_x_count):
                duration = random.uniform(0.3, 0.5)
                pyautogui.moveTo(c_start_x, c_start_y, duration=duration)
                pyautogui.click(c_start_x, c_start_y)
                time.sleep(0.1)
                pyautogui.click(c_start_x, c_start_y)
                time.sleep(2)

                # img_path = resource_path("axy.jpg")
                # img = cv2.imread(img_path)

                axy = find_text_position(resource_path("utils/axy.jpg"))
                bxy = find_text_position(resource_path("utils/bxy.jpg"))

                if (axy is not None and bxy is not None):
                    img = ImageGrab.grab(
                        bbox=(axy[0], axy[1], bxy[0] + 100, bxy[1] + 50))

                    name, level = baiduai(img)

                    # 保存截图
                    save_image(img, f'{file_path}/{level}',
                               f"{name}-{b+1}-{y+1}-{x+1}.png")

                else:
                    print("当前装备拦为空或者非稀有")
                c_start_x += x_grow

            c_start_x = start_x
            c_start_y += y_grow

        c_box_x += box_grow


get_window("暗黑破坏神IV")
time.sleep(3)

# x_grow = 51.5 x移动距离
# y_grow = 82   y移动距离
# start_x = 541 x起始坐标
# start_y = 411 y起始坐标
# x_count = 10  几列
# y_count = 5   几排
_run_()
