import pyautogui        # Mouse and window control
import win32gui         # finding the applications and bringing them to the Foreground
import win32con         # Used to manipulate the Application Windows if needed
import win32serviceutil  # Used to restart NX service
import time             # Cause who doesnt need time ?
import os

'''
Built based on screenshots from SonicWall NetExtender 8.6.263 on Windows OS. 

This script will automatically login and if it detects if the application is not running or the service is not running it will go and start them.

This script will need to be run with admin priviledges as it has to have access to the Services of Windows. 

Also, this only currently works when the Application is opened onto the main monitor.

version 1.0

'''

# pyautogui.PAUSE = .25
pyautogui.FAILSAFE = True
width, height = pyautogui.size()

# Global Variables
NXpath = "C:\\Program Files (x86)\\SonicWall\\SSL-VPN\\NetExtender\\NEGui.exe"
NXfront = False
NXuser = "test"
NXpass = "test"
NXdomain = "LocalDomain"

# This is the callback function for EnumWindows to get the list of Windows running


def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


# potentially need to use win32con to Maximize the window, only thing I do not like about it, is that it stretches it to the whole screen
top_windows = []
win32gui.EnumWindows(windowEnumerationHandler, top_windows)
for i in top_windows:
    if "netextender" in i[1].lower():
        print(i)
        NXfront = True
        win32gui.ShowWindow(i[0], 1)
        win32gui.SetForegroundWindow(i[0])
        break

if NXfront is False:
    NXsvcstatus = win32serviceutil.QueryServiceStatus(
        'SONICWALL_NetExtender')[1]
    if NXsvcstatus is 4:
        os.popen(NXpath)
    else:
        win32serviceutil.StartService("SONICWALL_NetExtender")
        os.popen(NXpath)

# waiting for the Lag of loading the GUI
time.sleep(2)

# Getting Window Size to reduce the region search
NXtitlebar = pyautogui.locateOnScreen('./images/NX_titlebar.png')
NXbottomright = pyautogui.locateOnScreen('./images/NXbottomright.png')

WindowSize = (NXtitlebar[0],
              NXtitlebar[1],
              NXbottomright[0] + NXbottomright[2] - NXtitlebar[0],
              NXbottomright[1] + NXbottomright[3] - NXtitlebar[1])

serverline = pyautogui.locateOnScreen(
    './images/serverline.png', region=WindowSize)

# Looking for the Server line to see if the program is loaded
# If the program is not loaded it will be looking for error messages and restarting service
while serverline is None:
    time.sleep(1)
    serverline = pyautogui.locateOnScreen(
        './images/images/serverline.png', region=WindowSize)
    errormsg = pyautogui.locateOnScreen(
        './images/NXError.png', region=WindowSize)

    if errormsg is not None:
        NXsvcerror = pyautogui.locateOnScreen(
            './images/NXsvcerror.png', region=WindowSize)
        if NXsvcerror is not None:
            win32serviceutil.StartService("SONICWALL_NetExtender")
            os.system("taskill /IM NEGui.exe")
            os.popen(NXpath)

serverline = pyautogui.locateOnScreen(
    './images/serverline.png', region=WindowSize)

pyautogui.click(serverline[0] + serverline[2] + 10,
                serverline[1] + serverline[3]/2)

clearbtn = pyautogui.locateOnScreen('./images/clearbtn.png', region=WindowSize)

pyautogui.click(clearbtn[0]+9, clearbtn[1]+12)

serverline = pyautogui.locateOnScreen(
    './images/serverline.png', region=WindowSize)
pyautogui.click(serverline[0] - serverline[2] + 10,
                serverline[1] - serverline[3])

emptysvrline = pyautogui.locateOnScreen(
    './images/emptysvrfield.png', region=WindowSize)

if emptysvrline is None:
    clearbtn = pyautogui.locateOnScreen(
        './images/clearbtn.png', region=WindowSize)
    try:
        pyautogui.click(clearbtn[0]+9, clearbtn[1]+12)
    except:
        pass

serverline = pyautogui.locateOnScreen(
    './images/serverline.png', region=WindowSize)
pyautogui.click(serverline[0] + serverline[2] + 10,
                serverline[1] + serverline[3]/2)

pyautogui.typewrite("192.168.227.1:4433")

unline = pyautogui.locateOnScreen('./images/username.png', region=WindowSize)
pyautogui.click(unline[0] + 10, unline[1] + 10)

unxbtn = pyautogui.locateOnScreen('./images/unxbtn.png', region=WindowSize)
if unxbtn is not None:
    pyautogui.click(unxbtn[0])

pyautogui.typewrite("test")
pyautogui.press('tab')
pyautogui.typewrite("test")
pyautogui.press('tab')
pyautogui.typewrite("LocalDomain")

connectbtn = pyautogui.locateOnScreen(
    './images/connectbtn.png', region=WindowSize)
pyautogui.click(connectbtn[0] + connectbtn[2]/2,
                connectbtn[1] + connectbtn[3]/2)

NXconnecting = pyautogui.locateOnScreen(
    './images/NXconnecting.png', region=WindowSize)
c = 0

# Will restart the program and service if there is an extended time stuck on Connecting
while NXconnecting is not None:
    time.sleep(1)
    NXconnecting = pyautogui.locateOnScreen(
        './images/NXconnecting.png', region=WindowSize)
    c += 1
    if c is 15:
        os.system("taskill /IM NEGui.exe")
        win32serviceutil.StartService("SONICWALL_NetExtender")
        os.popen(NXpath)

# Checking for an Error prompt
errormsg = pyautogui.locateOnScreen('./images/NXError.png', region=WindowSize)
NXloginfail = pyautogui.locateOnScreen(
    './images/NXLoginFailedMsg.png', region=WindowSize)

# this is checking for Login Failure
if NXloginfail is not None:
    NXreconbtn = pyautogui.locateOnScreen(
        './images/NXreconnectbtn.png', region=WindowSize)
    pyautogui.click(pyautogui.center(NXreconbtn))
    # Will need to re-try logging in from here
