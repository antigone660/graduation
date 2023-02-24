import win32gui
import win32api
import win32con
import win32ui
from PIL import Image
import time
import threading
import keyboard
import win32com.client as client
import pythoncom



def move_window_to_foreground(window_title):
    hwnd = win32gui.FindWindow(None, window_title)
    pythoncom.CoInitialize()
    shell = client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)


def click_mouse(window_title,x,y):
    hwnd = win32gui.FindWindow(None, window_title)
    move_window_to_foreground(window_title)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    x = left + x
    y = top + y
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

def capture_window(window_title, filename):
    hwnd = win32gui.FindWindow(None, window_title)
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top

    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, width, height)

    save_dc.SelectObject(bitmap)

        # 将窗口内容保存到位图中
    save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)

    # 将位图保存到文件中
    bitmap.SaveBitmapFile(save_dc, filename)

    # 释放资源
    win32gui.DeleteObject(bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)