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
from tkinter import scrolledtext, filedialog, simpledialog, ttk
import platform
import winsound
from pynput import keyboard

# 全局队列，用于在不同线程间安全通信
hotkey_queue = queue.Queue()

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
        self.master.title('QAutoCursor - 智能自动化 v1.0.2')
        self.master.geometry('500x720')
        self.master.configure(bg='#2E2E2E')
        self.master.resizable(False, False)

        # Style for TTK Notebook
        style = ttk.Style()
        style.theme_create("yaru", parent="alt", settings={
            "TNotebook": {"configure": {"tabmargins": [2, 5, 2, 0], "background": "#2E2E2E"}},
            "TNotebook.Tab": {
                "configure": {"padding": [10, 5], "background": "#3C3C3C", "foreground": "white"},
                "map": {"background": [("selected", "#BF3A3A")], "foreground": [("selected", "white")]}
            }
        })
        style.theme_use("yaru")

        self.targets = []
        self.selected_target_index = tk.IntVar(value=-1)
        self.target_list_frame = None

        self.action_mouse_var = tk.BooleanVar(value=True)
        self.mouse_action_type_var = tk.StringVar(value='左键单击')
        self.action_key_var = tk.BooleanVar(value=False)
        self.action_sound_var = tk.BooleanVar(value=True)
        self.stop_on_find_var = tk.BooleanVar(value=False)
        self.confidence_var = tk.DoubleVar(value=0.9)
        self.interval_var = tk.DoubleVar(value=5.0)
        self.monitor_thread = None
        self.sound_choice_var = tk.StringVar()
        self.available_sounds = []
        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()
        
        self.start_hotkey_var = tk.StringVar(value='Alt-1')
        self.stop_hotkey_var = tk.StringVar(value='Alt-2')
        
        self.hotkey_listener = None

        self.setup_ui()
        self.load_available_sounds()
        self.apply_hotkeys(initial_setup=True)
        self.master.after(100, self.process_log_queue)
        self.master.after(100, self.process_hotkey_queue)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.log("正在关闭应用...")
        if self.hotkey_listener and self.hotkey_listener.is_alive():
            self.hotkey_listener.stop()
        self.master.destroy()

    def parse_hotkey(self, key_string):
        """Converts a string like 'Control-F1' into a pynput-compatible format '<ctrl>+<f1>'."""
        parts = [p.strip().lower() for p in key_string.split('-')]
        
        modifiers = {
            'control': '<ctrl>', 'ctrl': '<ctrl>',
            'alt': '<alt>',
            'shift': '<shift>',
            'win': '<cmd>', 'command': '<cmd>'
        }
        
        pynput_keys = []
        for part in parts:
            if part in modifiers:
                pynput_keys.append(modifiers[part])
            # For pynput, function keys (f1, f2) and other special keys (home, end) need brackets.
            # Simple characters (a, b, c) do not. A good heuristic is length > 1.
            elif len(part) > 1:
                pynput_keys.append(f'<{part}>')
            else:
                pynput_keys.append(part)
        
        return '+'.join(pynput_keys)

    def apply_hotkeys(self, initial_setup=False):
        if self.hotkey_listener and self.hotkey_listener.is_alive():
            self.hotkey_listener.stop()

        start_key_str = self.start_hotkey_var.get()
        stop_key_str = self.stop_hotkey_var.get()
        
        if not start_key_str or not stop_key_str:
            self.log("错误: 快捷键配置不能为空。")
            return
            
        try:
            parsed_start = self.parse_hotkey(start_key_str)
            parsed_stop = self.parse_hotkey(stop_key_str)

            hotkeys = {
                parsed_start: self.safe_start_monitor,
                parsed_stop: self.safe_stop_monitor
            }

            self.hotkey_listener = keyboard.GlobalHotKeys(hotkeys)
            self.hotkey_listener.start()

            if hasattr(self, 'hotkey_info_label'):
                self.hotkey_info_label.config(text=f"快捷键: {start_key_str} 开始 | {stop_key_str} 停止")
            
            log_msg = "已初始化" if initial_setup else "已更新"
            self.log(f"全局快捷键{log_msg}: {start_key_str} 启动, {stop_key_str} 停止。")
        except Exception as e:
            self.log(f"错误: 设置快捷键失败 - {e}")

    def process_hotkey_queue(self):
        try:
            command = hotkey_queue.get_nowait()
            if command == 'start':
                self.start_monitor()
            elif command == 'stop':
                self.stop_monitor()
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self.process_hotkey_queue)

    def safe_start_monitor(self):
        hotkey_queue.put('start')

    def safe_stop_monitor(self):
        hotkey_queue.put('stop')

    def setup_ui(self):
        
        # --- Main Layout ---
        main_frame = tk.Frame(self.master, bg='#2E2E2E')
        main_frame.pack(fill='both', expand=True)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(pady=10, padx=10, fill='both', expand=True)

        tab_home = tk.Frame(notebook, bg='#2E2E2E')
        tab_settings = tk.Frame(notebook, bg='#2E2E2E')
        tab_log = tk.Frame(notebook, bg='#2E2E2E')

        notebook.add(tab_home, text='   主页   ')
        notebook.add(tab_settings, text=' 高级设置 ')
        notebook.add(tab_log, text='   日志   ')

        # --- Tab 1: Home ---
        self.setup_home_tab(tab_home)

        # --- Tab 2: Settings ---
        self.setup_settings_tab(tab_settings)
        
        # --- Tab 3: Log ---
        self.log_text = scrolledtext.ScrolledText(tab_log, wrap=tk.WORD, bg='#1E1E1E', fg='#D3D3D3', font=('Courier New', 10), relief=tk.FLAT)
        self.log_text.pack(fill='both', expand=True, padx=10, pady=10)
        self.log_text.config(state=tk.DISABLED)

    def setup_home_tab(self, parent):
        # --- Frame for target list ---
        list_container = tk.LabelFrame(parent, text='目标图像列表 (点击图像可设置操作点)', bg='#2E2E2E', fg='white', font=('Arial', 12, 'bold'), padx=10, pady=10)
        list_container.pack(pady=10, padx=10, fill='both', expand=True)

        canvas = tk.Canvas(list_container, bg='#3C3C3C', highlightthickness=0)
        scrollbar = tk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        self.target_list_frame = tk.Frame(canvas, bg='#3C3C3C')

        self.target_list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.target_list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # --- Frame for control buttons ---
        control_container = tk.Frame(parent, bg='#2E2E2E')
        control_container.pack(pady=10, padx=10, fill='x')

        # Image manipulation buttons
        image_btn_frame = tk.Frame(control_container, bg='#2E2E2E')
        image_btn_frame.pack(pady=5)
        tk.Button(image_btn_frame, text='导入图像', command=self.upload_image, bg='#1E1E1E', fg='white', relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(image_btn_frame, text='屏幕截图', command=self.snip_screen, bg='#1E1E1E', fg='white', relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        tk.Button(image_btn_frame, text='删除选中', command=self.remove_selected_target, bg='#BF3A3A', fg='white', relief=tk.FLAT).pack(side=tk.LEFT, padx=5)

        # --- 任务控制 ---
        task_control_frame = tk.LabelFrame(control_container, text='任务控制', bg='#2E2E2E', fg='white', font=('Arial', 12, 'bold'), padx=10, pady=10)
        task_control_frame.pack(pady=10, fill='x')
        
        button_container = tk.Frame(task_control_frame, bg='#2E2E2E')
        button_container.pack()
        
        self.start_btn = tk.Button(button_container, text='开始任务', bg='#1E1E1E', fg='white', font=('Arial', 12, 'bold'), relief=tk.FLAT, command=self.start_monitor, state=tk.NORMAL, padx=10, pady=5)
        self.start_btn.pack(side=tk.LEFT, padx=20)
        self.stop_btn = tk.Button(button_container, text='停止任务', bg='#BF3A3A', fg='white', font=('Arial', 12, 'bold'), relief=tk.FLAT, command=self.stop_monitor, state=tk.DISABLED, padx=10, pady=5)
        self.stop_btn.pack(side=tk.LEFT, padx=20)

        self.hotkey_info_label = tk.Label(task_control_frame, text="", bg='#2E2E2E', fg='grey', font=('Arial', 9))
        self.hotkey_info_label.pack(pady=4)

    def setup_settings_tab(self, parent):
        # --- 1. 自动化操作 ---
        action_frame = tk.LabelFrame(parent, text='自动化操作 (可组合)', bg='#2E2E2E', fg='white', font=('Arial', 12, 'bold'), padx=10, pady=10)
        action_frame.pack(pady=15, padx=20, fill='x')
        
        mouse_frame = tk.Frame(action_frame, bg='#2E2E2E')
        mouse_frame.pack(fill='x', pady=2)
        tk.Checkbutton(mouse_frame, text='鼠标操作  ', variable=self.action_mouse_var, bg='#2E2E2E', fg='white', selectcolor='#BF3A3A', command=self.toggle_mouse_input, activebackground='#2E2E2E', activeforeground='white').pack(side=tk.LEFT)
        mouse_options = ['左键单击', '右键单击', '双击', '移动至目标']
        self.mouse_option_menu = tk.OptionMenu(mouse_frame, self.mouse_action_type_var, *mouse_options)
        self.mouse_option_menu.config(bg='#3C3C3C', fg='white', activebackground='#555555', relief=tk.FLAT, highlightthickness=0)
        self.mouse_option_menu["menu"].config(bg='#3C3C3C', fg='white')
        self.mouse_option_menu.pack(side=tk.LEFT, padx=5)

        key_frame = tk.Frame(action_frame, bg='#2E2E2E')
        key_frame.pack(fill='x', pady=2)
        tk.Checkbutton(key_frame, text='键盘输入  ', variable=self.action_key_var, bg='#2E2E2E', fg='white', selectcolor='#BF3A3A', command=self.toggle_key_input, activebackground='#2E2E2E', activeforeground='white').pack(side=tk.LEFT)
        self.key_entry = tk.Entry(key_frame, font=('Arial', 10), bg='#3C3C3C', fg='white', relief=tk.FLAT, width=15, state=tk.DISABLED)
        self.key_entry.pack(side=tk.LEFT, padx=2)
        self.key_entry.insert(0, 'enter')

        sound_frame = tk.Frame(action_frame, bg='#2E2E2E')
        sound_frame.pack(fill='x', pady=2)
        tk.Checkbutton(sound_frame, text='播放提示音', variable=self.action_sound_var, bg='#2E2E2E', fg='white', selectcolor='#BF3A3A', activebackground='#2E2E2E', activeforeground='white', command=self.toggle_sound_input).pack(side=tk.LEFT)
        self.sound_option_menu = tk.OptionMenu(sound_frame, self.sound_choice_var, *['初始化...'])
        self.sound_option_menu.config(bg='#3C3C3C', fg='white', activebackground='#555555', relief=tk.FLAT, highlightthickness=0, width=20)
        self.sound_option_menu["menu"].config(bg='#3C3C3C', fg='white')
        self.sound_option_menu.pack(side=tk.LEFT, padx=5)
        self.toggle_sound_input()

        # --- 2. 任务设置 ---
        options_frame = tk.LabelFrame(parent, text='任务与快捷键', bg='#2E2E2E', fg='white', font=('Arial', 12, 'bold'), padx=10, pady=10)
        options_frame.pack(pady=15, padx=20, fill='x')
        
        stop_on_find_frame = tk.Frame(options_frame, bg='#2E2E2E')
        stop_on_find_frame.pack(anchor='w', fill='x')
        tk.Checkbutton(stop_on_find_frame, text='找到目标后停止任务', variable=self.stop_on_find_var, bg='#2E2E2E', fg='white', selectcolor='#BF3A3A', activebackground='#2E2E2E', activeforeground='white').pack(anchor='w')

        confidence_frame = tk.Frame(options_frame, bg='#2E2E2E')
        confidence_frame.pack(anchor='w', pady=5, fill='x')
        tk.Label(confidence_frame, text="图像相似度:", bg='#2E2E2E', fg='white').pack(side=tk.LEFT, padx=(0, 5))
        confidence_slider = tk.Scale(confidence_frame, from_=0.1, to=1.0, resolution=0.01, orient=tk.HORIZONTAL, variable=self.confidence_var, bg='#2E2E2E', fg='white', troughcolor='#3C3C3C', highlightthickness=0, length=250)
        confidence_slider.pack(side=tk.LEFT)

        interval_frame = tk.Frame(options_frame, bg='#2E2E2E')
        interval_frame.pack(anchor='w', pady=5, fill='x')
        tk.Label(interval_frame, text="检测间隔 (秒):", bg='#2E2E2E', fg='white').pack(side=tk.LEFT, padx=(0, 5))
        interval_entry = tk.Entry(interval_frame, textvariable=self.interval_var, font=('Arial', 10), bg='#3C3C3C', fg='white', relief=tk.FLAT, width=10)
        interval_entry.pack(side=tk.LEFT)

        hotkey_frame = tk.Frame(options_frame, bg='#2E2E2E')
        hotkey_frame.pack(anchor='w', pady=5)
        tk.Label(hotkey_frame, text="启动快捷键:", bg='#2E2E2E', fg='white').pack(side=tk.LEFT, padx=(0, 5))
        start_entry = tk.Entry(hotkey_frame, textvariable=self.start_hotkey_var, font=('Arial', 10), bg='#3C3C3C', fg='white', relief=tk.FLAT, width=10)
        start_entry.pack(side=tk.LEFT)
        tk.Label(hotkey_frame, text="停止快捷键:", bg='#2E2E2E', fg='white').pack(side=tk.LEFT, padx=(15, 5))
        stop_entry = tk.Entry(hotkey_frame, textvariable=self.stop_hotkey_var, font=('Arial', 10), bg='#3C3C3C', fg='white', relief=tk.FLAT, width=10)
        stop_entry.pack(side=tk.LEFT)
        tk.Button(hotkey_frame, text="应用", command=self.apply_hotkeys, bg='#1E1E1E', fg='white', relief=tk.FLAT).pack(side=tk.LEFT, padx=10)
        
    def log(self, message):
        self.log_queue.put(message)

    def process_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                
                if message == "TASK_COMPLETE":
                    self.stop_monitor()
                    continue
                if not hasattr(self, 'log_text'): return
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
            self.add_target_image(file_path)

    def remove_selected_target(self):
        index_to_remove = self.selected_target_index.get()
        if 0 <= index_to_remove < len(self.targets):
            self.targets.pop(index_to_remove)
            self.log(f"已删除目标 {index_to_remove + 1}。")
            self.update_target_list_ui()
        else:
            self.log("请先从列表中选择一个要删除的目标。")

    def update_target_list_ui(self):
        # Clear existing widgets
        for widget in self.target_list_frame.winfo_children():
            widget.destroy()

        if not self.targets:
            no_img_label = tk.Label(self.target_list_frame, text="请通过下方按钮添加目标图像", bg='#3C3C3C', fg='grey', font=('Arial', 10))
            no_img_label.pack(pady=20)
            return

        for i, target in enumerate(self.targets):
            item_frame = tk.Frame(self.target_list_frame, bg='#4A4A4A', bd=1, relief=tk.SOLID)
            item_frame.pack(fill='x', padx=5, pady=3)
            
            # --- Selection Radiobutton ---
            rb = tk.Radiobutton(item_frame, variable=self.selected_target_index, value=i, bg='#4A4A4A', activebackground='#4A4A4A', selectcolor='#BF3A3A')
            rb.pack(side=tk.LEFT, padx=5)
            
            # --- Thumbnail ---
            self.update_target_thumbnail(target) # Draw thumbnail with crosshair
            
            preview_label = tk.Label(item_frame, image=target['image_tk'], bg='#4A4A4A')
            preview_label.pack(side=tk.LEFT, padx=5, pady=5)
            preview_label.bind("<Button-1>", lambda e, index=i: self.on_preview_click(e, index))

            # --- Offset Info ---
            offset_text = f"操作点: {target['offset']}" if target['offset'] else "操作点: 中心"
            offset_label = tk.Label(item_frame, text=offset_text, bg='#4A4A4A', fg='white')
            offset_label.pack(side=tk.LEFT, padx=10)
        
        # Reset selection if it's now invalid
        if self.selected_target_index.get() >= len(self.targets):
            self.selected_target_index.set(-1)

    def on_preview_click(self, event, target_index):
        target = self.targets[target_index]
        if not target['image'] or not target['image_tk']:
            return

        thumb_w, thumb_h = target['image_tk'].width(), target['image_tk'].height()
        orig_w, orig_h = target['image'].size

        # For list items, the click event's x,y is relative to the label itself.
        click_x_on_thumb = event.x
        click_y_on_thumb = event.y
        
        if not (0 <= click_x_on_thumb < thumb_w and 0 <= click_y_on_thumb < thumb_h):
            return

        offset_x = click_x_on_thumb * (orig_w / thumb_w)
        offset_y = click_y_on_thumb * (orig_h / thumb_h)

        target['offset'] = (int(offset_x), int(offset_y))
        self.log(f"目标 {target_index + 1} 的操作点已更新为: {target['offset']}")
        self.update_target_list_ui() # Redraw to show the new crosshair

    def add_target_image(self, image_path_or_data):
        try:
            image = Image.open(image_path_or_data) if isinstance(image_path_or_data, str) else image_path_or_data
            new_target = {
                'image': image,
                'image_tk': None,
                'offset': None,
            }
            self.targets.append(new_target)

            self.log(f"已添加新目标，当前共 {len(self.targets)} 个目标。")
            self.update_target_list_ui()
        except Exception as e:
            self.log(f"错误: 无法加载图像 - {e}")

    def update_target_thumbnail(self, target):
        """Updates a target's thumbnail image including the crosshair."""
        MAX_THUMB_SIZE = (300, 60)
        thumb = target['image'].copy()
        thumb.thumbnail(MAX_THUMB_SIZE)
        
        draw = ImageDraw.Draw(thumb)
        scale_x = thumb.width / target['image'].width
        scale_y = thumb.height / target['image'].height

        if target['offset']:
            center_x = int(target['offset'][0] * scale_x)
            center_y = int(target['offset'][1] * scale_y)
        else: # Default to center
            center_x, center_y = thumb.width // 2, thumb.height // 2

        radius = 5
        draw.line((center_x - radius, center_y, center_x + radius, center_y), fill="red", width=2)
        draw.line((center_x, center_y - radius, center_x, center_y + radius), fill="red", width=2)

        target['image_tk'] = ImageTk.PhotoImage(thumb)

    def start_monitor(self):
        if self.start_btn['state'] == tk.DISABLED:
            self.log("任务已在运行中，请勿重复启动。")
            return

        active_targets_info = []
        for i, target in enumerate(self.targets):
            active_targets_info.append({
                'index': i,
                'image': target['image'],
                'offset': target['offset']
            })
        
        if not active_targets_info:
            self.log("错误: 请先添加至少一个目标图像。")
            return

        # Hide UI list
        self.target_list_frame.pack_forget()
        hidden_label = tk.Label(self.target_list_frame, text="\n...任务运行中...\n", bg='#3C3C3C', fg='grey', font=('Arial', 14))
        hidden_label.pack(fill='x', pady=50)

        action_mouse = self.action_mouse_var.get()
        mouse_action_type = self.mouse_action_type_var.get()
        action_key = self.action_key_var.get()
        action_sound = self.action_sound_var.get()
        key_to_press = self.key_entry.get()
        stop_on_find = self.stop_on_find_var.get()
        confidence_level = self.confidence_var.get()
        sound_choice = self.sound_choice_var.get()
        interval = self.interval_var.get()

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        self.stop_event.clear()
        self.monitor_thread = threading.Thread(
            target=self.monitor_loop,
            args=(action_mouse, mouse_action_type, action_key, key_to_press, action_sound, stop_on_find, confidence_level, sound_choice, interval, active_targets_info),
            daemon=True
        )
        self.monitor_thread.start()

    def stop_monitor(self):
        if self.stop_btn['state'] == tk.DISABLED:
            return # Task is not running, nothing to stop.

        self.stop_event.set()
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

        # Restore UI list
        for widget in self.target_list_frame.winfo_children():
            widget.destroy()
        self.target_list_frame.pack(fill='x')
        self.update_target_list_ui()
        
        self.log("任务已停止。")

    def monitor_loop(self, action_mouse, mouse_action_type, action_key, key_to_press, action_sound, stop_on_find, confidence_level, sound_choice, interval, active_targets):
        self.log(f"任务已启动，正在扫描屏幕 (相似度阈值: {confidence_level:.2f}, 间隔: {interval}s)...")
        while not self.stop_event.is_set():
            
            found_match_in_scan = False
            for target_info in active_targets:
                location_box = find_image_on_screen(target_info['image'], confidence=confidence_level)
                
                if location_box:
                    if target_info['offset']:
                        target_x = location_box.left + target_info['offset'][0]
                        target_y = location_box.top + target_info['offset'][1]
                    else: # Default to center
                        target_x = location_box.left + location_box.width / 2
                        target_y = location_box.top + location_box.height / 2
                    
                    target_point = (target_x, target_y)
                    self.log(f"已找到目标 {target_info['index'] + 1}，操作点: {target_point}")

                    if action_mouse:
                        if mouse_action_type == '左键单击':
                            pyautogui.click(target_point)
                            self.log("已执行: 鼠标左键单击。")
                        elif mouse_action_type == '右键单击':
                            pyautogui.rightClick(target_point)
                            self.log("已执行: 鼠标右键单击。")
                        elif mouse_action_type == '双击':
                            pyautogui.doubleClick(target_point)
                            self.log("已执行: 鼠标双击。")
                        elif mouse_action_type == '移动至目标':
                            pyautogui.moveTo(target_point, duration=0.3)
                            self.log("已执行: 移动鼠标至目标。")
                    
                    if action_sound and platform.system() == "Windows":
                        sound_to_play = sound_choice
                        if sound_to_play and sound_to_play != '无可用声音':
                            try:
                                flags = winsound.SND_ASYNC
                                if sound_to_play.lower().endswith('.wav'):
                                    flags |= winsound.SND_FILENAME
                                    self.log(f"已播放提示音: {sound_to_play}。")
                                else: # System sound
                                    flags |= winsound.SND_ALIAS
                                    self.log(f"已播放系统提示音: {sound_to_play}。")
                                winsound.PlaySound(sound_to_play, flags)
                            except Exception as e:
                                self.log(f"错误: 播放声音失败 - {e}")

                    if action_key:
                        pyautogui.press(key_to_press)
                        self.log(f"已执行: 键盘输入 '{key_to_press}'。")
                    
                    found_match_in_scan = True
                    break # Stop scanning other targets in this cycle
            
            if found_match_in_scan:
                if stop_on_find:
                    self.log("操作完成，任务已停止。")
                    self.log_queue.put("TASK_COMPLETE")
                    break
                else:
                    self.log("操作完成，继续扫描...")
                    time.sleep(interval)
            else:
                self.log(f"未在屏幕上找到任何目标(相似度>{confidence_level:.2f})，{interval}秒后重试...")
                time.sleep(interval)

    def load_available_sounds(self):
        """扫描并加载可用的声音文件和系统声音。"""
        if platform.system() != "Windows":
            self.available_sounds = ['无可用声音']
            self.sound_choice_var.set(self.available_sounds[0])
            self.update_sound_menu()
            return

        self.available_sounds = []
        
        # 扫描本地.wav文件
        try:
            local_wavs = [f for f in os.listdir('.') if f.lower().endswith('.wav')]
            if local_wavs:
                self.available_sounds.extend(sorted(local_wavs))
        except FileNotFoundError:
            self.log("警告: 无法扫描本地WAV文件目录。")
        
        if self.available_sounds:
            self.sound_choice_var.set(self.available_sounds[0])
        else:
            self.available_sounds = ['无可用声音']
            self.sound_choice_var.set(self.available_sounds[0])
        
        self.update_sound_menu()

    def update_sound_menu(self):
        """Helper function to update the OptionMenu."""
        if hasattr(self, 'sound_option_menu'):
            menu = self.sound_option_menu['menu']
            menu.delete(0, 'end')
            for sound in self.available_sounds:
                menu.add_command(label=sound, command=lambda value=sound: self.sound_choice_var.set(value))

    def toggle_mouse_input(self):
        state = tk.NORMAL if self.action_mouse_var.get() else tk.DISABLED
        self.mouse_option_menu.config(state=state)

    def toggle_key_input(self):
        if self.action_key_var.get():
            self.key_entry.config(state=tk.NORMAL)
        else:
            self.key_entry.config(state=tk.DISABLED)

    def toggle_sound_input(self):
        state = tk.NORMAL if self.action_sound_var.get() else tk.DISABLED
        if hasattr(self, 'sound_option_menu'):
            self.sound_option_menu.config(state=state)

    def snip_screen(self):
        self.master.withdraw()
        time.sleep(0.2)
        
        snipper = ScreenSnipper()
        region = snipper.get_region()
        
        if region:
            img = ImageGrab.grab(bbox=region)
            self.add_target_image(img)
        else:
            self.log("屏幕截图已取消。")
            
        self.master.deiconify()

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

def main():
    root = tk.Tk()
    app = ImageAutoClickerApp(root)
    root.mainloop()

if __name__ == '__main__':
    main() 