import win32gui
import win32con
import win32ui
import numpy as np
import cv2
import ctypes
import time
import os
import dxcam
from ctypes import windll
from pynput import keyboard

# 初始化 dxcam
camera = dxcam.create(output_idx=0)  # 默认第一个显示器
camera.start(target_fps=120)

# 常量定义
AIMLAB_WINDOW_NAME = "aimlab_tb"  
LOWER_CYAN = np.array([85, 120, 120])  # 青色 HSV 下界
UPPER_CYAN = np.array([95, 255, 255])  # 青色 HSV 上界
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
SMALL_TARGET_THRESHOLD = 900  # 目标大小阈值

# 全局变量
running = True

def on_press(key):
    """键盘监听，按 ESC 退出"""
    global running
    if key == keyboard.Key.esc:
        print("检测到 ESC，正在退出程序...")
        running = False
        return False

def get_dpi_scale():
    """获取系统 DPI 缩放比例"""
    try:
        user32 = windll.user32
        dpi = user32.GetDpiForSystem()
        print(f"系统 DPI: {dpi}")
        return dpi / 96.0
    except Exception as e:
        print(f"获取 DPI 失败: {e}, 使用默认值 1.0")
        return 1.0

DPI_ZOOM = get_dpi_scale()
print(f"系统 DPI 缩放: {DPI_ZOOM}")

def get_window_handle(title_section):
    """查找窗口句柄"""
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title_section.lower() in title.lower():
                windows.append((hwnd, title))
        return True
    
    windows = []
    win32gui.EnumWindows(callback, windows)
    if not windows:
        print(f"未找到标题包含 '{title_section}' 的窗口")
        return None
    hwnd, title = windows[0]
    print(f"找到窗口: {title} (句柄: {hwnd})")
    return hwnd

def get_window_rect(hwnd):
    """获取窗口位置和大小"""
    if not hwnd or not win32gui.IsWindow(hwnd):
        print("错误: 无效的窗口句柄")
        return None
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        print("错误: 窗口大小无效")
        return None
    print(f"窗口位置: ({left}, {top}), 大小: ({width}, {height})")
    return left, top, width, height




def capture_window(hwnd):
    """捕获窗口截图"""
    if not hwnd or not win32gui.IsWindow(hwnd):
        print("错误: 无效的窗口句柄或窗口已关闭")
        return None, None, None
    
    if win32gui.IsIconic(hwnd):
        print("窗口最小化，正在恢复...")
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.1)
    
    rect = get_window_rect(hwnd)
    if not rect:
        return None, None, None
    left, top, width, height = rect
    
    try:
        frame = camera.grab(region=(left, top, left + width, top + height))
        if frame is not None:
            img = np.array(frame)
            img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            cv2.imwrite("debug_dxcam.png", img)
            print(f"[DXCam] 截图保存为 debug_dxcam.png，大小: {img.shape}")
            return img, img_gray, rect
        else:
            raise RuntimeError("DXCam 返回空帧")
    except Exception as e:
        print(f"[DXCam 截图失败，回退 PrintWindow] {e}")
    
    #DXGI截图失败，则使用printwindow
    try:
        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
        save_dc.SelectObject(bitmap)
        
        result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 3)
        if not result:
            print("PrintWindow 调用失败")
            win32gui.DeleteObject(bitmap.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwnd_dc)
            return None, None, None
        
        bmp_str = bitmap.GetBitmapBits(True)
        img = np.frombuffer(bmp_str, dtype=np.uint8).reshape((height, width, 4))
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        cv2.imwrite("debug_screenshot.png", img)
        print(f"截图保存为 debug_screenshot.png，大小: {img.shape}")
        
        win32gui.DeleteObject(bitmap.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
        
        return img, img_gray, rect
    except Exception as e:
        print(f"PrintWindow 截图失败: {e}")
        return None, None, None


def template_img(template_img_path):
    """读取模板图像"""
    template_img = cv2.imread(template_img_path)
    if template_img is None:
        raise Exception(f"模板图像 {template_img_path} 读取失败")
    template_img_gray = cv2.cvtColor(template_img, cv2.COLOR_BGR2GRAY)
    return template_img_gray

def refine_position(image, center_x, center_y):
    """子像素精炼"""
    corners = np.array([[[center_x, center_y]]], dtype=np.float32)
    criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 30, 0.001)
    corners = cv2.cornerSubPix(image, corners, (5, 5), (-1, -1), criteria)
    return corners[0][0][0], corners[0][0][1]

def calculate_position(template_img_gray, screenshot_gray, width):
    """模板匹配查找目标位置"""
    scale_min, scale_max = 0.3, 1.2  # 优化小目标缩放范围
    max_val = 0.0
    best_center_x, best_center_y = None, None
    
    for scale in np.linspace(scale_min, scale_max, 30):
        resized_template = cv2.resize(template_img_gray, (0, 0), fx=scale, fy=scale)
        if resized_template.shape[0] > screenshot_gray.shape[0] or resized_template.shape[1] > screenshot_gray.shape[1]:
            continue
        
        result = cv2.matchTemplate(screenshot_gray, resized_template, cv2.TM_CCOEFF_NORMED)
        _, curr_max_val, _, max_loc = cv2.minMaxLoc(result)
        
        if curr_max_val >= 0.8 and curr_max_val > max_val:
            max_val = curr_max_val
            template_h, template_w = resized_template.shape[:2]   
            best_center_x = max_loc[0] + template_w // 2
            best_center_y = max_loc[1] + template_h // 2

            
            best_center_x, best_center_y = refine_position(screenshot_gray, best_center_x, best_center_y)
    
    if max_val >= 0.8:
        print(f"模板匹配成功，置信度: {max_val:.3f}, 中心坐标: ({best_center_x}, {best_center_y})")
        return best_center_x, best_center_y, max_val
    else:
        print(f"模板匹配失败，最高置信度: {max_val:.3f}")
        return None, None, None

def get_center_point(rect):
    """计算矩形中心点"""
    x, y, w, h = rect
    center_x = x + w // 2
    center_y = y + h // 2
    return center_x, center_y

# def move_mouse_relative_smooth(dx, dy, steps=50):
#     """平滑鼠标移动"""
#     try:
#         for i in range(steps):
#             ctypes.windll.user32.mouse_event(MOUSEEVENTF_MOVE, int(dx * DPI_ZOOM / steps), int(dy * DPI_ZOOM / steps), 0, 0)
#             time.sleep(0.01)
#         print(f"平滑鼠标移动: ({dx * DPI_ZOOM}, {dy * DPI_ZOOM})")
#     except Exception as e:
#         print(f"鼠标移动失败: {e}")
