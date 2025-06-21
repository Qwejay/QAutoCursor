import os
import sys
import time
from PIL import ImageGrab, Image, ImageTk, ImageDraw
import pytesseract
import cv2
import numpy as np
import pyautogui
import easyocr
import threading
import queue
import tkinter as tk
from tkinter import scrolledtext, filedialog, simpledialog
import platform
import winsound

# 自动检测tesseract路径
def auto_detect_tesseract():
    try:
        if sys.platform.startswith('win'):
            possible = [
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
            ]
            for path in possible:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    return True
            return False
        else: # mac/linux 默认在PATH
            return True
    except Exception:
        return False

# 智能截图（全屏）
def capture_screen():
    try:
        img_pil = ImageGrab.grab()
        # easyocr可以直接处理PIL图像或numpy数组，无需转为cv2格式
        return np.array(img_pil)
    except Exception as e:
        print(f"\n[ERROR] 截图失败: {e}")
        return None

# OCR识别 (easyocr)
def ocr_image(reader, img_np):
    try:
        # readtext返回 [[bbox, text, confidence], ...]
        return reader.readtext(img_np, detail=1)
    except Exception as e:
        print(f"\n[ERROR] OCR识别异常: {e}")
        return []

# 自动监控并点击（极简版, 使用easyocr）
def find_and_click_loop(reader, target_text):
    print(f"\n[*] 开始监控屏幕，寻找文本: '{target_text}'")
    print("[*] 按 Ctrl+C 可随时停止。")
    pyautogui.FAILSAFE = True
    
    while True:
        try:
            img = capture_screen()
            if img is None:
                time.sleep(2)
                continue
            
            results = ocr_image(reader, img)
            for (bbox, text, prob) in results:
                if target_text.lower() in text.lower():
                    # bbox是四个点的列表，取左上和右下计算中心点
                    (tl, tr, br, bl) = bbox
                    cx = int((tl[0] + br[0]) / 2)
                    cy = int((tl[1] + br[1]) / 2)
                    
                    print(f"\n[SUCCESS] 找到目标! 文本: '{text}', 坐标: ({cx},{cy}), 相似度: {prob:.2f}")
                    pyautogui.moveTo(cx, cy, duration=0.5, tween=pyautogui.easeInOutQuad)
                    pyautogui.click()
                    print("[*] 点击完成，任务结束。")
                    return
            
            print(".", end="", flush=True)
            time.sleep(1)

        except KeyboardInterrupt:
            print("\n\n[*] 用户中止操作。")
            break
        except Exception as e:
            print(f"\n[ERROR] 循环中发生未知错误: {e}")
            break

# --- 核心功能 ---
def capture_screen_region(region):
    """截取屏幕指定区域的图像。"""
    return np.array(ImageGrab.grab(bbox=region))

def find_image_on_screen(template_image, confidence=0.8):
    """在全屏上查找模板图像，返回Box(left, top, width, height)对象。"""
    try:
        location = pyautogui.locateOnScreen(template_image, confidence=confidence)
        return location
    except pyautogui.ImageNotFoundException:
        return None
    except Exception:
        return None

# --- UI部分 ---
class ImageAutoClickerApp:
    def __init__(self, master):
        self.master = master
        self.master.title('QAutoCursor 图像智能点击 v1.0.1')
        self.master.geometry('500x680')
        self.master.configure(bg='#2E2E2E')
        self.master.resizable(False, False)

        self.reference_image = None
        self.mouse_offset = None # (x, y) from top-left of reference image
        self.temp_ref_img_tk = None
        self.action_mouse_var = tk.BooleanVar(value=True)
        self.mouse_action_type_var = tk.StringVar(value='左键单击')
        self.action_key_var = tk.BooleanVar(value=False)
        self.action_sound_var = tk.BooleanVar(value=False)
        self.stop_on_find_var = tk.BooleanVar(value=True)
        self.monitor_thread = None
        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()

        self.setup_ui()
        self.master.after(100, self.process_log_queue)

    def setup_ui(self):
        # --- 1. 捕获目标 ---
        ref_frame = tk.Frame(self.master, bg='#2E2E2E', padx=10, pady=10)
        ref_frame.pack(pady=10, padx=20, fill='x')
        tk.Label(ref_frame, text='1. 捕获目标', bg='#2E2E2E', fg='white', font=('Arial', 12, 'bold')).pack(anchor='w')

        # 创建一个固定大小的框架来容纳预览图
        preview_frame = tk.Frame(ref_frame, bg='#3C3C3C', width=460, height=200)
        preview_frame.pack(pady=5)
        preview_frame.pack_propagate(False) # 防止子控件改变框架大小

        self.ref_image_label = tk.Label(preview_frame, text='暂无参考图', bg='#3C3C3C', fg='grey')
        self.ref_image_label.pack(expand=True, fill='both')
        
        btn_container = tk.Frame(ref_frame, bg='#2E2E2E')
        btn_container.pack()
        tk.Button(btn_container, text='上传图片', command=self.upload_image, bg='#1E1E1E', fg='white', relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_container, text='截取区域', command=self.snip_screen, bg='#1E1E1E', fg='white', relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        self.anchor_btn = tk.Button(btn_container, text='设置定位点', command=self.set_anchor_point, bg='#1E1E1E', fg='white', relief=tk.FLAT, state=tk.DISABLED)
        self.anchor_btn.pack(side=tk.LEFT, padx=5)

        # --- 2. 响应动作 ---
        action_frame = tk.Frame(self.master, bg='#2E2E2E', padx=10, pady=10)
        action_frame.pack(pady=10, padx=20, fill='x')
        tk.Label(action_frame, text='2. 响应动作 (可多选)', bg='#2E2E2E', fg='white', font=('Arial', 12, 'bold')).pack(anchor='w')
        
        # Mouse Action Row
        mouse_frame = tk.Frame(action_frame, bg='#2E2E2E')
        mouse_frame.pack(fill='x', pady=2)
        tk.Checkbutton(mouse_frame, text='鼠标动作  ', variable=self.action_mouse_var, bg='#2E2E2E', fg='white', selectcolor='#BF3A3A', command=self.toggle_mouse_input, activebackground='#2E2E2E', activeforeground='white').pack(side=tk.LEFT)
        
        mouse_options = ['左键单击', '右键单击', '双击', '移动到此处']
        self.mouse_option_menu = tk.OptionMenu(mouse_frame, self.mouse_action_type_var, *mouse_options)
        self.mouse_option_menu.config(bg='#3C3C3C', fg='white', activebackground='#555555', relief=tk.FLAT, highlightthickness=0)
        self.mouse_option_menu["menu"].config(bg='#3C3C3C', fg='white')
        self.mouse_option_menu.pack(side=tk.LEFT)

        # Key Action Row
        key_frame = tk.Frame(action_frame, bg='#2E2E2E')
        key_frame.pack(fill='x', pady=2)
        tk.Checkbutton(key_frame, text='键盘响应  ', variable=self.action_key_var, bg='#2E2E2E', fg='white', selectcolor='#BF3A3A', command=self.toggle_key_input, activebackground='#2E2E2E', activeforeground='white').pack(side=tk.LEFT)
        self.key_entry = tk.Entry(key_frame, font=('Arial', 10), bg='#3C3C3C', fg='white', relief=tk.FLAT, width=15, state=tk.DISABLED)
        self.key_entry.pack(side=tk.LEFT)
        self.key_entry.insert(0, 'enter')

        # Sound Action Row
        sound_frame = tk.Frame(action_frame, bg='#2E2E2E')
        sound_frame.pack(fill='x', pady=2)
        tk.Checkbutton(sound_frame, text='声音提示 (仅Windows)', variable=self.action_sound_var, bg='#2E2E2E', fg='white', selectcolor='#BF3A3A', activebackground='#2E2E2E', activeforeground='white').pack(side=tk.LEFT)

        # --- 3. 任务配置 ---
        options_frame = tk.Frame(self.master, bg='#2E2E2E', padx=10, pady=10)
        options_frame.pack(pady=10, padx=20, fill='x')
        tk.Label(options_frame, text='3. 任务配置', bg='#2E2E2E', fg='white', font=('Arial', 12, 'bold')).pack(anchor='w')
        tk.Checkbutton(options_frame, text='找到匹配后停止监控', variable=self.stop_on_find_var, bg='#2E2E2E', fg='white', selectcolor='#BF3A3A', activebackground='#2E2E2E', activeforeground='white').pack(anchor='w')

        # --- 4. 控制与日志 ---
        control_frame = tk.Frame(self.master, bg='#2E2E2E')
        control_frame.pack(pady=10)
        self.start_btn = tk.Button(control_frame, text='开始监控', bg='#1E1E1E', fg='white', font=('Arial', 12, 'bold'), relief=tk.FLAT, command=self.start_monitor, state=tk.NORMAL, padx=10, pady=5)
        self.start_btn.pack(side=tk.LEFT, padx=10)
        self.stop_btn = tk.Button(control_frame, text='停止监控', bg='#BF3A3A', fg='white', font=('Arial', 12, 'bold'), relief=tk.FLAT, command=self.stop_monitor, state=tk.DISABLED, padx=10, pady=5)
        self.stop_btn.pack(side=tk.LEFT, padx=10)
        
        log_frame = tk.Frame(self.master, bg='#2E2E2E')
        log_frame.pack(pady=10, padx=20, fill='both', expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, bg='#1E1E1E', fg='#D3D3D3', font=('Courier New', 10), relief=tk.FLAT)
        self.log_text.pack(fill='both', expand=True)
        self.log_text.config(state=tk.DISABLED)

    def log(self, message):
        self.log_queue.put(message)

    def process_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                
                if message == "TASK_COMPLETE":
                    self.stop_monitor()
                    continue

                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.master.after(100, self.process_log_queue)

    def upload_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp")])
        if file_path:
            self.set_reference_image(file_path)

    def set_reference_image(self, image_path_or_data):
        self.reference_image = Image.open(image_path_or_data) if isinstance(image_path_or_data, str) else image_path_or_data
        self.mouse_offset = None # Reset offset on new image
        self.update_preview()
        self.anchor_btn.config(state=tk.NORMAL)
        self.log("参考图已设置。可通过'设置定位点'指定鼠标位置（默认为中心）。")

    def update_preview(self):
        if not self.reference_image:
            self.ref_image_label.config(image='', text="暂无参考图")
            return
        
        MAX_PREVIEW_SIZE = (460, 200)
        thumb = self.reference_image.copy()
        thumb.thumbnail(MAX_PREVIEW_SIZE)
        
        draw = ImageDraw.Draw(thumb)
        scale_x = thumb.width / self.reference_image.width
        scale_y = thumb.height / self.reference_image.height
        
        if self.mouse_offset:
            center_x, center_y = int(self.mouse_offset[0] * scale_x), int(self.mouse_offset[1] * scale_y)
        else: # Default to center
            center_x, center_y = thumb.width // 2, thumb.height // 2

        radius = 8
        draw.line((center_x - radius, center_y, center_x + radius, center_y), fill="red", width=2)
        draw.line((center_x, center_y - radius, center_x, center_y + radius), fill="red", width=2)

        self.ref_img_tk = ImageTk.PhotoImage(thumb)
        self.ref_image_label.config(image=self.ref_img_tk, text="")

    def snip_screen(self):
        self.master.withdraw()
        time.sleep(0.2) # 等待主窗口隐藏
        
        snipper = ScreenSnipper()
        region = snipper.get_region()
        
        if region:
            img = ImageGrab.grab(bbox=region)
            self.set_reference_image(img)
        else:
            self.log("截图已取消。")
            
        self.master.deiconify()

    def toggle_mouse_input(self):
        state = tk.NORMAL if self.action_mouse_var.get() else tk.DISABLED
        self.mouse_option_menu.config(state=state)

    def toggle_key_input(self):
        if self.action_key_var.get():
            self.key_entry.config(state=tk.NORMAL)
        else:
            self.key_entry.config(state=tk.DISABLED)

    def start_monitor(self):
        if not self.reference_image:
            self.log("错误: 请先提供一个参考图！")
            return

        # 隐藏预览图，防止自我检测
        if hasattr(self, 'ref_img_tk'):
            self.temp_ref_img_tk = self.ref_img_tk
        self.ref_image_label.config(image="", text="...监控运行中...")
        self.master.update_idletasks() # 强制UI立即更新

        action_mouse = self.action_mouse_var.get()
        mouse_action_type = self.mouse_action_type_var.get()
        action_key = self.action_key_var.get()
        action_sound = self.action_sound_var.get()
        key_to_press = self.key_entry.get()
        stop_on_find = self.stop_on_find_var.get()

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        self.stop_event.clear()
        self.monitor_thread = threading.Thread(
            target=self.monitor_loop,
            args=(action_mouse, mouse_action_type, action_key, key_to_press, action_sound, stop_on_find),
            daemon=True
        )
        self.monitor_thread.start()

    def stop_monitor(self):
        self.stop_event.set()
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

        # 恢复预览图
        if self.temp_ref_img_tk:
            self.ref_image_label.config(image=self.temp_ref_img_tk)
            self.temp_ref_img_tk = None # 清理临时存储
        else:
            self.ref_image_label.config(image="", text="暂无参考图")
        
        self.log("监控已停止。")

    def monitor_loop(self, action_mouse, mouse_action_type, action_key, key_to_press, action_sound, stop_on_find):
        self.log("开始监控...")
        while not self.stop_event.is_set():
            location_box = find_image_on_screen(self.reference_image)
            
            if location_box:
                if self.mouse_offset:
                    target_x = location_box.left + self.mouse_offset[0]
                    target_y = location_box.top + self.mouse_offset[1]
                else: # Default to center
                    target_x = location_box.left + location_box.width / 2
                    target_y = location_box.top + location_box.height / 2
                
                target_point = (target_x, target_y)
                self.log(f"找到匹配图像，目标点: {target_point}")

                if action_mouse:
                    if mouse_action_type == '左键单击':
                        pyautogui.click(target_point)
                        self.log("执行鼠标左键单击。")
                    elif mouse_action_type == '右键单击':
                        pyautogui.rightClick(target_point)
                        self.log("执行鼠标右键单击。")
                    elif mouse_action_type == '双击':
                        pyautogui.doubleClick(target_point)
                        self.log("执行鼠标双击。")
                    elif mouse_action_type == '移动到此处':
                        pyautogui.moveTo(target_point, duration=0.3)
                        self.log("执行鼠标移动。")
                
                if action_sound and platform.system() == "Windows":
                    winsound.Beep(1000, 200) # 1000Hz frequency for 200ms
                    self.log("执行声音提示。")

                if action_key:
                    pyautogui.press(key_to_press)
                    self.log(f"执行键盘响应: 按下 '{key_to_press}'")
                
                if stop_on_find:
                    self.log("任务完成，监控已停止。")
                    self.log_queue.put("TASK_COMPLETE")
                    break
                else:
                    self.log("操作完成，继续监控...")
                    time.sleep(1) # Pause after action before next scan
            else:
                self.log("未找到匹配图像，1秒后重试...")
                time.sleep(1)

    def set_anchor_point(self):
        if not self.reference_image:
            self.log("错误: 请先提供参考图！")
            return

        def callback(offset):
            if offset:
                self.mouse_offset = offset
                self.log(f"定位点已更新为: {offset}")
                self.update_preview()
        
        AnchorSetter(self.master, self.reference_image, callback)

# --- 屏幕截图工具 ---
class ScreenSnipper:
    def __init__(self):
        self.root = tk.Toplevel()
        self.root.attributes("-fullscreen", True)
        self.root.attributes("-alpha", 0.2)
        self.root.overrideredirect(True)
        self.canvas = tk.Canvas(self.root, cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.region = None

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def on_button_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

    def on_mouse_drag(self, event):
        cur_x, cur_y = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_button_release(self, event):
        end_x, end_y = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        x1, y1 = min(self.start_x, end_x), min(self.start_y, end_y)
        x2, y2 = max(self.start_x, end_x), max(self.start_y, end_y)
        self.region = (x1, y1, x2, y2)
        self.root.quit()
        self.root.destroy()

    def get_region(self):
        self.root.mainloop()
        return self.region

class AnchorSetter:
    def __init__(self, master, image, callback):
        self.root = tk.Toplevel(master)
        self.root.title("设置定位点 (单击图片选择)")
        self.image = image
        self.callback = callback

        self.tk_image = ImageTk.PhotoImage(image)
        self.canvas = tk.Canvas(self.root, width=image.width, height=image.height, cursor="cross")
        self.canvas.pack()
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_image)
        
        self.canvas.bind("<Button-1>", self.on_click)
        
        self.root.transient(master)
        self.root.grab_set()
        self.root.wait_window(self.root)

    def on_click(self, event):
        offset = (event.x, event.y)
        self.root.destroy()
        self.callback(offset)

def main():
    root = tk.Tk()
    app = ImageAutoClickerApp(root)
    root.mainloop()

if __name__ == '__main__':
    main() 