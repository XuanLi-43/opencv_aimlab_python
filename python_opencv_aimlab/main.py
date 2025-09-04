from screen_and_mouse import *
from pynput import keyboard

def main():
    print("主程序正在运行，按 ESC 键退出")
    
    template_path = r"./BlueBall.png"
    template_img_gray = template_img(template_path)
    
    cv2.namedWindow("this", cv2.WINDOW_FREERATIO)
    cv2.moveWindow("this", 0, 0)
    cv2.resizeWindow("this", int(1280 / DPI_ZOOM), int(720 / DPI_ZOOM))
    
    while running:
        start_time = time.time()
        
        hwnd = get_window_handle(AIMLAB_WINDOW_NAME)
        if not hwnd:
            time.sleep(1)
            continue
        
        frame, screenshot_gray, rect = capture_window(hwnd)
        if frame is None or frame.size == 0:
            print("截图失败")
            cv2.imshow("this", frame)
            time.sleep(0.1)
            continue
        
        left, top, width, height = rect
        window_center_x = width // 2
        window_center_y = height // 2
        
        hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
        binary = cv2.inRange(hsv, LOWER_CYAN, UPPER_CYAN)
        kernel = np.ones((3, 3), np.uint8)
        binary = cv2.dilate(binary, kernel, iterations=1)
        binary = cv2.erode(binary, kernel, iterations=1)
        cv2.imwrite("debug_binary.png", binary)
        print("二值化图像保存为 debug_binary.png")
        
        contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        print(f"检测到轮廓数量: {len(contours)}")
        
        if not contours:
            print("[警告] HSV 未找到青色目标")
            cv2.imshow("this", frame)
            if cv2.waitKey(1) & 0xFF == 27:
                break
            time.sleep(0.02)
            continue
        
        min_distance = float('inf')
        candidate_x, candidate_y = None, None
        candidate_size = 0
        candidate_rect = None
        
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            if w < 15 or h < 15:
                continue
            
            center_x, center_y = get_center_point((x, y, w, h))
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
            cv2.circle(frame, (center_x, center_y), 8, (0, 0, 255), -1)
            
            dx = center_x - window_center_x
            dy = center_y - window_center_y
            distance = dx * dx + dy * dy
            
            if distance < min_distance:
                min_distance = distance
                candidate_x, candidate_y = center_x, center_y
                candidate_size = w * h
                candidate_rect = (x, y, w, h)
            
            info = f"{dx} {dy}"
            cv2.putText(frame, info, (center_x, center_y), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)
        
        if candidate_x is None or candidate_y is None:
            print("[警告] 未找到有效候选目标")
            
            mouse_left = 30
            ctypes.windll.user32.mouse_event(MOUSEEVENTF_MOVE, int(mouse_left * DPI_ZOOM), 0, 0, 0)
            print("未检测到目标，向左移动")
            
            cv2.imshow("this", frame)
            if cv2.waitKey(1) & 0xFF == 27:
                break
            time.sleep(0.02)
            continue
        
        final_x, final_y = candidate_x, candidate_y
        if candidate_size < SMALL_TARGET_THRESHOLD:
            print("HSV 检测目标太小，使用模板匹配")
            crop_x = max(0, candidate_rect[0] - 150)
            crop_y = max(0, candidate_rect[1] - 150)
            crop_w = min(width, candidate_rect[0] + candidate_rect[2] + 150) - crop_x
            crop_h = min(height, candidate_rect[1] + candidate_rect[3] + 150) - crop_y
            crop_gray = screenshot_gray[crop_y:crop_y + crop_h, crop_x:crop_x + crop_w]
            cv2.imwrite("debug_crop.png", crop_gray)
            
            target_x, target_y, max_val = calculate_position(template_img_gray, crop_gray, crop_w)
            
            if max_val is not None and max_val >= 0.8:
                final_x = crop_x + target_x
                final_y = crop_y + target_y
                print("使用模板匹配坐标")
            else:
                print("模板匹配失败，使用 HSV 坐标")
        else:
            print("HSV 检测目标正常，使用 HSV 坐标")
        
        try:
            
    # final_x, final_y 是窗口内坐标，加上窗口左上角得到全局屏幕坐标
            global_x = left + final_x
            global_y = top + final_y

                # 窗口中心的全局坐标
            window_center_global_x = left + window_center_x
            window_center_global_y = top + window_center_y

                # 鼠标偏移量
            dx = global_x - window_center_global_x
            dy = global_y - window_center_global_y
            distance = dx * dx + dy * dy

            print(f"最终目标坐标(窗口内): ({final_x}, {final_y}), "
                f"全局坐标: ({global_x}, {global_y}), 偏移: ({dx}, {dy}), 距离: {distance:.2f}")
                

                # 在窗口画面里标记目标
            cv2.drawMarker(frame, (int(final_x), int(final_y)), (0, 255, 0), cv2.MARKER_CROSS, 20, 2)

                # move_mouse_relative_smooth(dx, dy, steps=20)
                # 直接一次性移动鼠标
            ctypes.windll.user32.mouse_event(MOUSEEVENTF_MOVE, int(dx * DPI_ZOOM), int(dy * DPI_ZOOM), 0, 0)

                # 如果目标足够近则点击
            if distance < 500:
                ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    
                ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                print("[动作完成] 已点击目标")
            
        except Exception as e:
                print(f"鼠标操作失败: {e}")
        
        elapsed = (time.time() - start_time) * 1000
        cv2.putText(frame, f"{elapsed:.2f}ms", (100, 100), cv2.FONT_HERSHEY_SIMPLEX, 3, (0, 0, 255), 2)
        
        cv2.imshow("this", frame)
        print("显示图像")
        if cv2.waitKey(1) & 0xFF == 27:
            break
        
        time.sleep(0.02)
    
    cv2.destroyAllWindows()
    print("程序结束")

if __name__ == "__main__":
    time.sleep(5)
    print("5秒后启动")
    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    main()
    print("程序结束")