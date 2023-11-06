import cv2
import datetime
import numpy
import pywinauto
import time
import win32ui
import win32gui
from ctypes import windll
from multiprocessing import Process
from pickle import FALSE
from PIL import Image
from pynput import keyboard
from pywinauto.application import Application

toplist, winlist = [], []


#################################################### Help Functions ####################################################### 

def CalcWindow(hwnd):
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    w = right - left
    h = bot - top
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    saveDC.SelectObject(saveBitMap)
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    image = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    return image

def ClickLeft(window, coords):
    if coords != 0:
        window.click(button='left', coords=coords)

def ClickObject(window, screen, images):
    for image in images:
        template_image = cv2.imread(image)
        result = cv2.matchTemplate(screen, template_image, cv2.TM_CCOEFF_NORMED) 
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        x, y = max_loc
        coords = (x,y)
        MouseMove(window, coords)
        ClickLeft(window, coords)   

def EnumCb(hwnd, results):
    winlist.append((hwnd, win32gui.GetWindowText(hwnd)))

def GenshinQuestClicker():
    win32gui.EnumWindows(EnumCb, toplist)
    chrome = [(hwnd, title) for hwnd, title in winlist if 'genshin impact' in title.lower()]
    app = Application()
    app.connect(handle=chrome[0][0])

    ############ Genshin populates two windows. Sometimes you have to access the second one. #############

    ############ Acces first window #########

    hwnd = chrome[0][0]

    ############ Acces second window #########

    # hwnd = chrome[1][0]

    window = GetWindowHandle(app, hwnd)
    rect = pywinauto.win32structures.RECT()
    pywinauto.win32functions.GetWindowRect(window, rect)
    while True:
        if ReadFromFile("stop.txt") == "0": 
            ClickObject(window, numpy.array(CalcWindow(hwnd)), ["genshin.png", "genshin2.png"] )
        time.sleep(0.1)  

def GetWindowHandle(app, playerHwnd):
    for window in app.windows():
        if playerHwnd == window.handle:
            return window
    return None

def MouseMove(window, coords):
    window.click_input(button='move', coords=coords, button_down=False, button_up=False, key_down=False, key_up=False)    

def Now():
    return datetime.datetime.now()

def On_press(key):
    if any([key in z for z in [{keyboard.Key.shift}]]):
        if ReadFromFile("stop.txt") == "0":
            WriteToFile("stop.txt", "1")
        else:
            WriteToFile("stop.txt", "0")  
    
def ReadFromFile(fileName):
    with open(fileName) as file:
        return file.read()

def WriteToFile(fileName, input):
    with open(fileName, "w") as file:
        file.write(str(input))
       
if __name__ == '__main__':
    time.sleep(1)
    processListener = keyboard.Listener(on_press=On_press)
    processClicl = Process(target = GenshinQuestClicker)
    processClicl.start()
    processListener.start()
    processListener.join()
    processClicl.join()