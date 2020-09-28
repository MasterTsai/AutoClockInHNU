# -*- coding: utf-8 -*-
"""
Created on Thu Feb 20 23:07:37 2020

@author: 
"""

"""
1. 需下载chromedriver。
2. 需下载tesseract，用于识别验证码。
"""

from selenium import webdriver
import re  # 用于正则
from PIL import Image  # 用于打开图片和对图片处理
import pytesseract  # 用于图片转文字
from time import sleep
from datetime import datetime
from datetime import date
from smtplib import SMTP_SSL
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from win32 import win32api, win32gui, win32print
from win32.win32api import GetSystemMetrics
from win32.lib import win32con
from random import random


## 随机体温 36°~37°
def rd_body_temp():
    yesterday_temp = 36 + round(random(), 1)
    today_temp = 36 + round(random(), 1)
    return yesterday_temp, today_temp


def import_msg(info):
    info_size = info.shape[0]
    msg = ""
    for i in range(info_size):
        msg += str(i+1) + '.' + info.loc[i, "姓名"] + '\n'

    return msg


# 获取屏幕分辨率
def get_real_resolution():
    """获取真实的分辨率"""
    hDC = win32gui.GetDC(0)
    # 横向分辨率
    w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    return w, h


def get_screen_size():
    """获取缩放后的分辨率"""
    w = GetSystemMetrics(0)
    h = GetSystemMetrics(1)
    return w, h


def get_screen_scale_rate():
    real_resolution = get_real_resolution()
    screen_size = get_screen_size()
    screen_scale_rate = round(real_resolution[0] / screen_size[0], 2)
    #print('缩放比例为：%f' % screen_scale_rate)
    return screen_scale_rate


# 获取个人信息
def get_info():
    personal_data = pd.read_excel("./个人信息.xlsx")
    return personal_data


# 获取当下时间 打卡成功 记录日志
def clock_in_successfully(temperature):
    datetime_now = datetime.now()
    today_log = str(datetime_now) + " ...... " + your_name + "第" + str(t + 1) + "次打卡成功"
    yesterday_temp = '昨日体温：' + str(temperature[0])
    today_temp = '今日体温：' + str(temperature[1])
    today_log += '\n' + yesterday_temp
    today_log += '\n' + today_temp
    print(yesterday_temp)
    print(today_temp)
    return today_log


# 获取验证码坐标
def get_range():
    ele_loc = driver.find_element_by_xpath('//*[@id="app"]/div/div[3]/div[3]/img').location
    ele_size = driver.find_element_by_xpath('//*[@id="app"]/div/div[3]/div[3]/img').size
    rate = get_screen_scale_rate()
    left = (ele_loc['x'] + 2) * rate
    top = (ele_loc['y'] + 2) * rate
    right = (ele_loc['x'] + ele_size['width'] - 2) * rate
    bottom = (ele_loc['y'] + ele_size['height'] - 2) * rate
    myrange = (left, top, right, bottom)
    #print(myrange)
    return myrange


# 获取验证码
def get_vcode():
    # 获取截图
    driver.save_screenshot(str(i + 1) + 'pictures.png')  # 全屏截图
    page_snap_obj = Image.open(str(i + 1) + 'pictures.png')
    myrange = get_range()
    pic_crop = page_snap_obj.crop(myrange)
    # pic_crop.show()
    return pic_crop


def processing_image():
    pic_crop = get_vcode()  # 获取验证码
    img = pic_crop.convert("L")  # 转灰度
    pixdata = img.load()
    w, h = img.size
    threshold = 185  # 该阈值不适合所有验证码，具体阈值请根据验证码情况设置
    # 遍历所有像素，大于阈值的为黑色
    for y in range(h):
        for x in range(w):
            if pixdata[x, y] < threshold:
                pixdata[x, y] = 0
            else:
                pixdata[x, y] = 255
    return img


def delete_spot():
    img = processing_image()
    data = img.getdata()
    w, h = img.size
    black_point = 0
    for x in range(1, w - 1):
        for y in range(1, h - 1):
            mid_pixel = data[w * y + x]  # 中央像素点像素值
            if mid_pixel < 50:  # 找出上下左右四个方向像素点像素值
                top_pixel = data[w * (y - 1) + x]
                left_pixel = data[w * y + (x - 1)]
                down_pixel = data[w * (y + 1) + x]
                right_pixel = data[w * y + (x + 1)]
                # 判断上下左右的黑色像素点总个数
                if top_pixel < 10:
                    black_point += 1
                if left_pixel < 10:
                    black_point += 1
                if down_pixel < 10:
                    black_point += 1
                if right_pixel < 10:
                    black_point += 1
                if black_point < 1:
                    img.putpixel((x, y), 255)
                black_point = 0
    # img.show()
    return img


def image_str():
    image = delete_spot()
    # 需设置tesseract路径
    pytesseract.pytesseract.tesseract_cmd = per_data.loc[0, "tesseract路径"]
    result = pytesseract.image_to_string(image)  # 图片转文字
    resultj = re.sub(u"([^\u4e00-\u9fa5\u0030-\u0039\u0041-\u005a\u0061-\u007a])", "", result)  # 去除识别出来的特殊字符
    result_four = resultj[0:4]  # 只获取前4个字符
    #print(resultj)  # 打印识别的验证码
    return result_four


def img_name():
    global date_today
    date_today = str(date.today())
    screenshot_name = date_today + your_name + '打卡.png'
    return date_today, screenshot_name


## 今日是否打卡##
def loc(mydate):
    
    date = int(mydate)
    today = datetime.now().weekday()
    # 周日是第一天
    if today == 6:
        weekday = 1
    # 周六是第七天
    else:
        weekday = today + 2
    # 判断是第几个星期
    location = int((date-weekday)/7 + 1) * 7 + weekday
    return location
    
            
def if_clock_in_today():
    date = date_today[8:]
    try: 
        driver.find_elements_by_xpath(
        '//*[@id="app"]/div/div[2]/div[2]/div/div[1]/div/div/div/div/section/div/div[3]/div['
        + str(loc(date)) + ']/div')[0].click()
        return 1
    except:
        return 0


#######################################
###########Send Emails#################
#######################################


def send_main():
    file_path = screenshot_name
    send_message(file_path)


def send_message(file_path):
    # 服务器 pyinstaller -F email.py
    host_server = per_data.loc[ppl, "服务器"]

    # 授权码
    pwd = per_data.loc[ppl, "发件邮箱授权码"]

    # 发件人邮箱
    from_email = per_data.loc[ppl, "发件邮箱"]

    # email address of receiver
    to_email = per_data.loc[ppl, "收件邮箱"]

    # 邮件正文
    msg = MIMEMultipart()

    # 1. 邮件头
    # 编码集
    # 万国码 utf-8 全世界通用
    msg['Subject'] = Header(subject, 'utf-8')

    # 2. 发件箱
    msg['From'] = from_email

    # 3. 收件箱
    msg['To'] = to_email

    # 需要渲染
    msg.attach(MIMEText(date_today + '打卡成功', 'html', 'utf-8'))

    # 文件路径 picture.jpg 关联
    # 添加附件 ssl
    image = MIMEText(open(screenshot_name, 'rb').read(), 'html', 'utf-8')
    # txt jpg png mp4
    image['Content-Type'] = 'image/png'
    image.add_header('Content-Disposition', 'attachment', filename=screenshot_name)
    msg.attach(image)

    # 登录
    smtp = SMTP_SSL(host_server)
    smtp.login(from_email, pwd)
    smtp.sendmail(from_email, to_email, msg.as_string())

    print(your_name, '发送完毕-->' + to_email)
    smtp.quit()


def send_summ(admin, mssg):
    host_server = per_data.loc[admin, "服务器"]
    pwd = per_data.loc[admin, "发件邮箱授权码"]
    from_email = per_data.loc[admin, "发件邮箱"]
    to_email = per_data.loc[admin, "收件邮箱"]

    # 邮件正文
    msg = MIMEMultipart()

    msg['Subject'] = Header(date_today+'打卡报告', 'utf-8')

    msg['From'] = from_email

    msg['To'] = to_email

    msg.attach(MIMEText(mssg))

    # 登录
    smtp = SMTP_SSL(host_server)
    smtp.login(from_email, pwd)
    smtp.sendmail(from_email, to_email, msg.as_string())

    print('汇总信息发送完毕-->' + to_email)
    smtp.quit()


per_data = get_info()
num_ppl = per_data.shape[0]
print('请输入对应的编号！')
print(import_msg(per_data))
start = input('从哪位同学开始？\n')
start = int(start)
admin = input('汇总信息发给谁？\n')
admin = int(admin) - 1 


# 初始化日志
today_log = []

url = "http://fangkong.hnu.edu.cn/app/#/login"

for ppl in range(start-1, num_ppl):
    ####################################
    # 个人信息
    # 姓名
    your_name = per_data.loc[ppl, "姓名"]
    # 你的学号
    your_id = per_data.loc[ppl, "学号"]
    # 你的密码
    your_password = per_data.loc[ppl, "密码"]
    # 省份
    province_index = per_data.loc[ppl, "省份"]
    # 城市
    city_index = per_data.loc[ppl, "城市"]
    # 镇/区
    district_index = per_data.loc[ppl, "镇区"]
    # 详细地址
    your_address = per_data.loc[ppl, "地址"]
    # executable_path
    executable_path = per_data.loc[ppl, "chromedriver路径"]
    # 体温
    temperature = rd_body_temp()
    ######################################
    ######################################
    # 邮件设置
    # date 截图命名
    date_today, screenshot_name = img_name()
    mydate = date_today[8:]
    # 主题
    subject = date_today + your_name + "打卡截图"
    ######################################
    ######################################
    # 开始打卡
    for t in range(5):
        flag = 0
        try:
            print(your_name, '第'+str(t+1)+'次打卡')
            # executable_path是chromedriver.exe(自行下载)的位置
            # options 设置
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
            chrome_options.add_argument(
                'user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Safari/537.36"'
            )
            driver = webdriver.Chrome(executable_path=executable_path, options=chrome_options)

            driver.maximize_window()
            driver.get(url)

            sleep(1)

            # 自动输入学号和密码
            driver.find_elements_by_xpath('//*[@id="app"]/div/div[3]/div[1]/input')[0].send_keys(your_id)
            driver.find_elements_by_xpath('//*[@id="app"]/div/div[3]/div[2]/input')[0].send_keys(your_password)
            # driver.find_elements_by_xpath('//*[@id="app"]/div/div[3]/button')[0].click()
            sleep(1)

            for i in range(10):
                try:
                    # 获取验证码 处理验证码
                    print('第%i次尝试' % (i + 1))
                    driver.find_elements_by_xpath('//*[@id="app"]/div/div[3]/div[4]')[0].click()

                    sleep(1)

                    img = processing_image()
                    # img.show()
                    vcode = image_str()
                    print(vcode)

                    # 输入验证码
                    driver.find_elements_by_xpath('//*[@id="app"]/div/div[3]/div[3]/div/input')[0].send_keys(vcode)
                    #sleep(1)
                    # 点击确认
                    driver.find_elements_by_xpath('//*[@id="app"]/div/div[3]/button')[0].click()
                    sleep(1)

                    # 进入第二个页面
                    driver.find_elements_by_xpath(
                        '//*[@id="app"]/div/div[2]/div[2]/div/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div[2]/input')[
                        0].click()
                    print('第%i次尝试成功' % (i + 1))
                    break

                except:
                    # 清空对话框
                    driver.find_element_by_xpath('//*[@id="app"]/div/div[3]/div[3]/div/input').clear()
            if i == 10:
                print("10次尝试都失败，请您手动登录！")
            
            sleep(1)
            # 填写体温
            driver.find_elements_by_xpath(
                '//*[@id="app"]/div/div[2]/div[2]/div/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div[2]/input')[0].send_keys(str(temperature[0]))

            driver.find_elements_by_xpath(
                '//*[@id="app"]/div/div[2]/div[2]/div/div[1]/div/div/div/div[2]/div[2]/div/div[3]/div[2]/input')[0].send_keys(str(temperature[1]))


            driver.find_elements_by_xpath(
                '//*[@id="app"]/div/div[2]/div[2]/div/div[1]/div/div/div/div[2]/button')[0].click()
            sleep(2)
            driver.find_elements_by_xpath(
                '//*[@id="app"]/div/div[2]/div[2]/div/div[1]/div/div/div/div/section/div/div[3]/div['
                + str(loc(mydate)) + ']/div')[0].click()

            # 打卡成功
            flag = 1
            temp = clock_in_successfully(temperature)
            print(temp)
            today_log.append(temp)
            driver.save_screenshot(screenshot_name)
            sleep(0.5)
            # 退出
            driver.quit()
            # 打卡完毕
            print(your_name, '第'+str(t+1)+'次打卡成功')
            # 发邮件
            send_main()
            break

        except:
            if_ci = if_clock_in_today()
            if if_ci == 1:
                print(your_name + "今日已打卡！")
                driver.quit()
                break
            else:
                print(your_name, '第'+str(t+1)+'次打卡失败')
                print("垃圾玩意儿！")
                if t+1 == 5:
                    print(your_name, '5次打卡都失败，请手动打卡！')
                   
        else:
            continue
        if if_ci == 1:
            break


temp = ''

if len(today_log) != 0:
    for i in range(len(today_log)):
        temp += today_log[i] + '\n'
    today_log = temp
    with open("./打卡日志.txt", "a") as f:
        f.write('\n'+today_log)
        f.close()
    
    send_summ(admin, today_log)
