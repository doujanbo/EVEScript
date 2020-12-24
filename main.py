import cv2
import numpy
import aircv as ac
import os
import sys
import time
from ctypes import *
import win32com
import win32api
import win32com.client
import json
from PIL import ImageGrab
import random
import pyttsx3


def matchImg(imgsrc,imgobj,num,confidencevalue=0.5):
    imsrc = ac.imread(imgsrc)
    imobj = ac.imread(imgobj)

    match_result = ac.find_template(imsrc,imobj,num)
    if match_result is not None:
        match_result['shape']=(imsrc.shape[1],imsrc.shape[0])
    return match_result

def bmpsave(bmpname):
    im = ImageGrab.grab()
    im.save('./Screenshot/' + bmpname + '.bmp','bmp')

#进程内注册插件,模块所在的路径按照实际位置修改
#hkmdll = windll.LoadLibrary("C:/Users/douxiaoyao/PycharmProjects/pythonProject/mouse/wyhkm.dll")

#创建对象
#try:
wyhkm=win32com.client.Dispatch("wyp.hkm")
#except:
#print("创建对象失败!")
#sys.exit(0)
#获得模块版本号
version=wyhkm.GetVersion()
print("无涯键鼠盒子模块版本："+hex(version))
#查找设备,这个只是例子,参数中的VID和PID要改成实际值

DevId=wyhkm.SearchDevice(0x2612, 0x1701, 0)

if DevId==-1:
    print("未找到无涯键鼠盒子")
    sys.exit(0)
#打开设备
if not wyhkm.Open(DevId):
    print("打开无涯键鼠盒子失败")
    sys.exit(0)
#打开资源管理器快捷键Win+E

coolrdinate = (0, 0)
LockCoordinate = (0, 0)
OreCoolrdinate = (0, 0)

coordinate = (matchImg("./Screenshot/firstslide.bmp","./Screenshot/Asteroid.bmp", 0.9)['result'])

LockCoordinate = (matchImg("./Screenshot/firstslide.bmp","./Screenshot/locking.bmp", 0.9)['result'])

def MoveToAdd(x,y):
    time.sleep(random.randint(66, 100) * 0.001)
    x = x + random.randint(-2,3)
    y = y + random.randint(-2,3)
    wyhkm.MoveTo(x, y)
    time.sleep(random.randint(66, 100) * 0.001)

def LockAsteroid():
    global coordinate
    global  LockCoordinate

    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(coordinate[0], coordinate[1])  # 移动到小行星的位置
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(coordinate[0], coordinate[1] + 16)  # 移动到小行星的位置
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(coordinate[0], coordinate[1] + 32)  # 移动到小行星的位置
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(coordinate[0], coordinate[1] + 48)  # 移动到小行星的位置
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(coordinate[0], coordinate[1] + 64)  # 移动到小行星的位置
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(coordinate[0], coordinate[1] + 80)  # 移动到小行星的位置
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

engine = pyttsx3.init()

def sayhello(strhello):
    pass
    '''
    engine.say(strhello)
    engine.runAndWait()
'''

# 注意，没有本句话是没有声音的


sayhello('你好')
time.sleep(5)
sayhello('开始锁定')
LockAsteroid()
time.sleep(5)
sayhello('刷新截图')
bmpsave('RefreshScreenshot')


def F1F2():
    global coolrdinate
    time.sleep(1)
    bmpsave('RefreshScreenshot')
    try:
        coolrdinate  = matchImg("./Screenshot/RefreshScreenshot.bmp","./Screenshot/MineGun.bmp", 0.99)['result']
    except:
        print('没有空闲的矿枪')
        return
    MoveToAdd(coolrdinate[0], coolrdinate[1])
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftClick()
    bmpsave('RefreshScreenshot')
    time.sleep(1)


def OreTransfer():
    global OreCoolrdinate
    try:
        OreCoolrdinate = matchImg("./Screenshot/RefreshScreenshot.bmp","./Screenshot/Oreinthewarehouse.bmp", 0.9)['result']
    except:
        print('仓库里没有矿石')
        return

    MoveToAdd(OreCoolrdinate[0], OreCoolrdinate[1])
    wyhkm.LeftClick()
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.KeyDown('A')
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.KeyUP('Ctrl')
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.KeyUP('A')
    time.sleep(random.randint(66, 100) * 0.001)
    OreCoolrdinate = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/whale.bmp", 0.9)['result']
    wyhkm.LeftDown()
    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(OreCoolrdinate[0],OreCoolrdinate[1])
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftUP()
    time.sleep(random.randint(66, 100) * 0.001)
def duidie():
    xyxy = (0, 0)
    bmpsave('RefreshScreenshot')
    try:
        wyhkm.LeftUp()
        bmpsave('RefreshScreenshot')
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/jdjk.bmp", 0.9)['result']
        MoveToAdd(xyxy[0], xyxy[1])
        wyhkm.LeftClick()
        time.sleep(0.5)
        bmpsave('RefreshScreenshot')
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/yukuang.bmp", 0.9)['result']
        MoveToAdd(xyxy[0], xyxy[1])
        time.sleep(random.randint(66, 100) * 0.001)
        wyhkm.RightClick()
        time.sleep(random.randint(66, 100) * 0.001)
        bmpsave('RefreshScreenshot')
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/duidie.bmp", 0.9)['result']
        MoveToAdd(xyxy[0], xyxy[1])
        time.sleep(random.randint(66, 100) * 0.001)
        sayhello("堆叠")
        wyhkm.LeftClick()
        time.sleep(3)
    except:
        return

def transport(): #移动矿石
    xyxy = (0, 0)
    duidie()
    try:
        wyhkm.LeftUp()
        bmpsave('RefreshScreenshot')
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/jdjk.bmp", 0.9)['result']
        MoveToAdd(xyxy[0], xyxy[1])
        wyhkm.LeftClick()
        time.sleep(0.5)
        bmpsave('RefreshScreenshot')
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp","./Screenshot/yukuang.bmp", 0.9)['result']
        MoveToAdd(xyxy[0],xyxy[1])
        wyhkm.LeftDown()
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/ksc.bmp", 0.9)['result']
        MoveToAdd(xyxy[0], xyxy[1])
        wyhkm.LeftUp()
        time.sleep(random.randint(66, 100) * 0.001)
        wyhkm.LeftClick()
    except:
        sayhello('小鱼舰队机库没有矿石')
        print('小鱼舰队机库没有矿石')
        LittleWhale('jipin')
        return
    pass

def LittleWhale(name):   #切换界面
    xyxy = (0, 0)
    wyhkm.KeyDown('Alt')
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.KeyDown('Tab')
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.KeyUP('Tab')
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.KeyUP('Alt')
    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(1440, 900)
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.LeftClick()
    time.sleep(random.randint(66, 100) * 0.001)
    MoveToAdd(416,880)
    time.sleep(random.randint(66, 100) * 0.001+1)
    bmpsave('RefreshScreenshot')
    try:
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/" + name + ".bmp", 0.9)['result']
    except:
        return
    MoveToAdd(xyxy[0]- 30, xyxy[1] + 36)
    time.sleep(random.randint(66, 100) * 0.001 + 0.5)
    wyhkm.LeftClick()
    time.sleep(random.randint(66, 100) * 0.001)

def chekbmp(bmpname):
    bmpsave('RefreshScreenshot')
    while True:
        try:
            matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/" + bmpname + ".bmp", 0.9)
            sayhello("发现了")
            return
        except:
            sayhello("没发现")
            bmpsave('RefreshScreenshot')
            pass

def jinzhan():
    xyxy = (0, 0)
    while True:
        try:
            MoveToAdd(620, 450)
            wyhkm.RightClick()
            time.sleep(random.randint(66, 100) * 0.001)
            time.sleep(1)
            bmpsave('RefreshScreenshot')
            # 移到建筑
            sayhello("移到建筑")
            xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/jianzhu.bmp", 0.9)['result']
            MoveToAdd(xyxy[0], xyxy[1])
            time.sleep(0.5)
            bmpsave('RefreshScreenshot')
            # 移到铁壁
            sayhello("移到铁壁")
            xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/tiebi.bmp", 0.9)['result']
            MoveToAdd(xyxy[0], xyxy[1])
            time.sleep(0.6)
            # 停靠
            sayhello("停靠")
            bmpsave('RefreshScreenshot')
            xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/tingkao.bmp", 0.5)['result']
            MoveToAdd(xyxy[0], xyxy[1])
            wyhkm.LeftClick()
            time.sleep(random.randint(66, 100) * 0.001)
            sayhello("还没到空间站")
            time.sleep(5)
        except:
            sayhello("已经进站了")
            return


def full():
    xyxy = (0,0)
    bmpsave('RefreshScreenshot')
    sayhello("检测是否满了")
    try:

        matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/full.bmp", 0.9)['result']
        sayhello("回站")
        wyhkm.KeyPress("Enter")
        MoveToAdd(620, 450)
        wyhkm.RightClick()
        time.sleep(random.randint(66, 100) * 0.001)
        time.sleep(1)
        bmpsave('RefreshScreenshot')
        #移到建筑
        sayhello("移到建筑")
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/jianzhu.bmp", 0.9)['result']
        MoveToAdd(xyxy[0], xyxy[1])
        time.sleep(0.5)
        bmpsave('RefreshScreenshot')
        #移到铁壁
        sayhello("移到铁壁")
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/tiebi.bmp", 0.9)['result']
        MoveToAdd(xyxy[0], xyxy[1])
        time.sleep(0.6)
        #停靠
        sayhello("停靠,停止加成")
        wyhkm.KeyPress('F1')
        time.sleep(random.randint(66, 100) * 0.001)
        wyhkm.KeyPress('F2')
        time.sleep(random.randint(66, 100) * 0.001)
        wyhkm.KeyPress('F3')
        bmpsave('RefreshScreenshot')
        xyxy = matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/tingkao.bmp", 0.5)['result']
        MoveToAdd(xyxy[0], xyxy[1])
        wyhkm.LeftClick()
        time.sleep(random.randint(66, 100) * 0.001)
        #进站

    except:
        sayhello("没满")
        pass



#time.sleep(3)
#transport()
#full()
#jinzhan()


while True:
    iii = 0
    if matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/OrePhotos.bmp", 0.9):
        sayhello('直接就发现了矿石')
        print("直接就发现了矿石1")
        time.sleep(2)
        F1F2()
        time.sleep(1)
        bmpsave('RefreshScreenshot')
    else:
        while matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/OrePhotos.bmp", 0.9) == None:
            sayhello('直接就发现了矿石')
            print('直接就发现了矿石2')
            LockAsteroid()
            time.sleep(5)
            bmpsave('RefreshScreenshot')
        sayhello('等待了一会才发现了矿石')
        print('等待了一会才发现了矿石3')
        time.sleep(1)
        F1F2()
        time.sleep(2)

    if bool(matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/TwoMineQuns.bmp", 0.9)):
        while True:
            sayhello('正在挖矿')
            print('正在挖矿4')
            if iii == 6:
                iii = 0
                LittleWhale('dou')
                time.sleep(0.5)
                transport()
                time.sleep(0.5)
                LittleWhale('jipin')
            iii = iii + 1
            F1F2()
            F1F2()
            if matchImg("./Screenshot/RefreshScreenshot.bmp", "./Screenshot/TwoMineQuns.bmp", 0.9) == None:
                bmpsave('RefreshScreenshot')
                break
            OreTransfer()
            time.sleep(0.5)
            time.sleep(2)

#关闭设备
wyhkm.Close()
input()


