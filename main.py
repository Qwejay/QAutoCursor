import os
import json
import time
import traceback
import tempfile
from PIL import ImageGrab, Image, ImageTk, ImageDraw
import pyautogui
import threading
import queue
import tkinter as tk
from tkinter import scrolledtext, filedialog, ttk
from playsound import playsound
from pynput import keyboard
import sys
import winsound

hotkey_queue = queue.Queue()

def sound_dingding():
    winsound.Beep(1000, 120)
    winsound.Beep(1500, 120)

def sound_up():
    winsound.Beep(800, 80)
    winsound.Beep(1200, 80)
    winsound.Beep(1600, 80)

def sound_down():
    winsound.Beep(1600, 80)
    winsound.Beep(1200, 80)
    winsound.Beep(800, 80)

def sound_error():
    winsound.Beep(400, 200)
    winsound.Beep(300, 300)

def sound_success():
    winsound.Beep(1200, 100)
    winsound.Beep(1800, 200)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- UI部分 ---
class ImageAutoClickerApp:
    def __init__(self, master):
        self.log_queue = queue.Queue()  # 必须最先初始化
        self.master = master
        self.master.title('QAutoCursor - 智能自动化 v1.1.0')

        try:
            icon_path = resource_path("icon.ico")
            if os.path.exists(icon_path):
                # 推荐用iconphoto，兼容性更好
                img = None
                try:
                    img = tk.PhotoImage(file=icon_path)
                except Exception:
                    pass
                if img:
                    self.master.iconphoto(True, img)
        except Exception as e:
            print(f"警告: 无法加载图标 - {e}")

        # 极简风格配色（全部白底）
        MAIN_BG = '#FFFFFF'  # 纯白
        MAIN_FG = '#222222'  # 深灰黑
        SUB_FG = '#888888'   # 浅灰
        BTN_BG = '#222222'   # 深灰黑
        BTN_FG = '#FFFFFF'   # 按钮字体色
        BTN_BG_HOVER = '#444444'  # 按钮悬停色
        TAB_SELECTED_BG = '#222222'
        TAB_SELECTED_FG = '#FFFFFF'
        TAB_BG = '#FFFFFF'   # Tab未选中也为白
        TAB_FG = '#222222'
        ENTRY_BG = '#FFFFFF' # 输入框白底
        ENTRY_FG = '#222222'
        BORDER_COLOR = '#E0E0E0'
        DEL_BTN_BG = '#E53935'  # 低饱和度红
        DEL_BTN_FG = '#FFFFFF'
        DEL_BTN_BG_HOVER = '#B71C1C'
        LOG_BG = '#FFFFFF'   # 日志区白底
        LOG_FG = '#888888'
        # 保存为成员变量，后续用
        self._theme = dict(MAIN_BG=MAIN_BG, MAIN_FG=MAIN_FG, SUB_FG=SUB_FG, BTN_BG=BTN_BG, BTN_FG=BTN_FG, BTN_BG_HOVER=BTN_BG_HOVER,
                           TAB_SELECTED_BG=TAB_SELECTED_BG, TAB_SELECTED_FG=TAB_SELECTED_FG, TAB_BG=TAB_BG, TAB_FG=TAB_FG,
                           ENTRY_BG=ENTRY_BG, ENTRY_FG=ENTRY_FG, BORDER_COLOR=BORDER_COLOR, DEL_BTN_BG=DEL_BTN_BG, DEL_BTN_FG=DEL_BTN_FG,
                           DEL_BTN_BG_HOVER=DEL_BTN_BG_HOVER, LOG_BG=LOG_BG, LOG_FG=LOG_FG)

        self.master.geometry('500x720')
        self.master.configure(bg=MAIN_BG)
        self.master.resizable(False, False)

        # Style for TTK Notebook
        style = ttk.Style()
        style.theme_create("minimal", parent="alt", settings={
            "TNotebook": {"configure": {"tabmargins": [2, 5, 2, 0], "background": MAIN_BG, "borderwidth": 0}},
            "TNotebook.Tab": {
                "configure": {"padding": [10, 5], "background": TAB_BG, "foreground": TAB_FG, "font": ('Segoe UI', 10, 'bold')},
                "map": {"background": [("selected", TAB_SELECTED_BG)], "foreground": [("selected", TAB_SELECTED_FG)]}
            },
            # 圆角按钮风格
            "Rounded.TButton": {
                "configure": {
                    "font": ('Segoe UI', 10, 'bold'),
                    "foreground": BTN_FG,
                    "background": BTN_BG,
                    "borderwidth": 0,
                    "padding": [12, 6],
                    "relief": "flat"
                },
                "map": {
                    "background": [("active", BTN_BG_HOVER), ("!active", BTN_BG)],
                    "foreground": [("disabled", '#AAAAAA'), ("!disabled", BTN_FG)]
                }
            },
            "Danger.TButton": {
                "configure": {
                    "font": ('Segoe UI', 10, 'bold'),
                    "foreground": DEL_BTN_FG,
                    "background": DEL_BTN_BG,
                    "borderwidth": 0,
                    "padding": [12, 6],
                    "relief": "flat"
                },
                "map": {
                    "background": [("active", DEL_BTN_BG_HOVER), ("!active", DEL_BTN_BG)],
                    "foreground": [("disabled", '#FFFFFF'), ("!disabled", DEL_BTN_FG)]
                }
            }
        })
        style.theme_use("minimal")

        try:
            style.element_create("RoundedButton", "from", "clam")
            style.layout("Rounded.TButton", [
                ("RoundedButton", None, {
                    "children": [
                        ("Button.focus", {"children": [
                            ("Button.padding", {"children": [
                                ("Button.label", {"side": "left", "expand": 1})
                            ]})
                        ]})
                    ]
                })
            ])
        except Exception:
            pass

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
        self.sound_style_var = tk.StringVar(value='叮叮')  # 新增，默认叮叮
        self.stop_event = threading.Event()
        
        self.start_hotkey_var = tk.StringVar(value='Alt-1')
        self.stop_hotkey_var = tk.StringVar(value='Alt-2')
        
        self.hotkey_listener = None
        self.config_file = "config.json"
        self.targets_dir = "targets"
        self.sound_data = {} # To hold sound file paths
        self.temp_sound_files = [] # To clean up on exit

        self.setup_ui()
        self.load_available_sounds() # Load sounds after setting up UI
        self.load_config() # Load config after UI is setup
        self.apply_hotkeys(initial_setup=True)
        self.master.after(100, self.process_log_queue)
        self.master.after(100, self.process_hotkey_queue)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.log("正在保存配置并关闭应用...")
        self.save_config()
        self.stop_event.set()

        # Clean up temporary sound files
        for temp_path in self.temp_sound_files:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    self.log(f"信息: 已清理临时声音文件 {os.path.basename(temp_path)}")
            except Exception as e:
                self.log(f"警告: 清理临时文件失败 {temp_path} - {e}")

        if self.hotkey_listener and self.hotkey_listener.is_alive():
            self.hotkey_listener.stop()
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=1.0)
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

    def find_image_on_screen(self, template_image, confidence=0.8):
        """在全屏上查找模板图像，返回Box(left, top, width, height)对象。"""
        try:
            location = pyautogui.locateOnScreen(template_image, confidence=confidence)
            return location
        except pyautogui.ImageNotFoundException:
            return None
        except Exception as e:
            self.log(f"查找图像时发生未知错误: {e}")
            return None

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
        theme = self._theme
        # --- Main Layout ---
        main_frame = tk.Frame(self.master, bg=theme['MAIN_BG'])
        main_frame.pack(fill='both', expand=True)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(pady=10, padx=10, fill='both', expand=True)

        tab_home = tk.Frame(notebook, bg=theme['MAIN_BG'])
        tab_settings = tk.Frame(notebook, bg=theme['MAIN_BG'])
        tab_log = tk.Frame(notebook, bg=theme['MAIN_BG'])

        notebook.add(tab_home, text='   主页   ')
        notebook.add(tab_settings, text=' 高级设置 ')
        notebook.add(tab_log, text='   日志   ')

        # --- Tab 1: Home ---
        self.setup_home_tab(tab_home)
        # --- Tab 2: Settings ---
        self.setup_settings_tab(tab_settings)
        # --- Tab 3: Log ---
        self.log_text = scrolledtext.ScrolledText(tab_log, wrap=tk.WORD, bg=theme['LOG_BG'], fg=theme['LOG_FG'], font=('Segoe UI', 10), relief=tk.FLAT)
        self.log_text.pack(fill='both', expand=True, padx=10, pady=10)
        self.log_text.config(state=tk.DISABLED)
        self.log_text.config(bg=theme['MAIN_BG'])

    def setup_home_tab(self, parent):
        theme = self._theme
        # --- Frame for target list ---
        list_container = tk.LabelFrame(parent, text='目标图像列表 (点击图像可设置鼠标点击位置)', bg=theme['MAIN_BG'], fg=theme['MAIN_FG'], font=('Segoe UI', 12, 'bold'), padx=10, pady=10, bd=1, relief=tk.GROOVE)
        list_container.pack(pady=10, padx=10, fill='both', expand=True)

        canvas = tk.Canvas(list_container, bg=theme['MAIN_BG'], highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        self.target_list_frame = tk.Frame(canvas, bg=theme['MAIN_BG'])

        self.target_list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.target_list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # --- Frame for control buttons ---
        control_container = tk.Frame(parent, bg=theme['MAIN_BG'])
        control_container.pack(pady=10, padx=10, fill='x')

        # Image manipulation buttons
        image_btn_frame = tk.Frame(control_container, bg=theme['MAIN_BG'])
        image_btn_frame.pack(pady=5)
        ttk.Button(image_btn_frame, text='导入图像', command=self.upload_image, style='Rounded.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(image_btn_frame, text='屏幕截图', command=self.snip_screen, style='Rounded.TButton').pack(side=tk.LEFT, padx=5)

        # --- 任务控制 ---
        task_control_frame = tk.LabelFrame(control_container, text='任务控制', bg=theme['MAIN_BG'], fg=theme['MAIN_FG'], font=('Segoe UI', 12, 'bold'), padx=10, pady=10, bd=1, relief=tk.GROOVE)
        task_control_frame.pack(pady=10, fill='x')
        button_container = tk.Frame(task_control_frame, bg=theme['MAIN_BG'])
        button_container.pack()
        self.start_btn = ttk.Button(button_container, text='开始任务', command=self.start_monitor, style='Rounded.TButton')
        self.start_btn.pack(side=tk.LEFT, padx=20)
        self.stop_btn = ttk.Button(button_container, text='停止任务', command=self.stop_monitor, style='Danger.TButton')
        self.stop_btn.pack(side=tk.LEFT, padx=20)
        self.hotkey_info_label = tk.Label(task_control_frame, text="", bg=theme['MAIN_BG'], fg='#666666', font=('Segoe UI', 9))
        self.hotkey_info_label.pack(pady=4)

    def setup_settings_tab(self, parent):
        theme = self._theme
        # --- 1. 自动化操作 ---
        action_frame = tk.LabelFrame(parent, text='自动化操作 (可组合)', bg=theme['MAIN_BG'], fg=theme['MAIN_FG'], font=('Segoe UI', 12, 'bold'), padx=10, pady=10, bd=1, relief=tk.GROOVE)
        action_frame.pack(pady=15, padx=20, fill='x')
        mouse_frame = tk.Frame(action_frame, bg=theme['MAIN_BG'])
        mouse_frame.pack(fill='x', pady=2)
        tk.Checkbutton(mouse_frame, text='鼠标操作  ', variable=self.action_mouse_var,
            bg=theme['MAIN_BG'], fg=theme['MAIN_FG'],
            selectcolor=theme['MAIN_BG'],
            activebackground=theme['MAIN_BG'], activeforeground=theme['MAIN_FG'],
            highlightthickness=0, bd=0,
            command=self.toggle_mouse_input).pack(side=tk.LEFT)
        mouse_options = ['左键单击', '右键单击', '双击', '移动至目标']
        self.mouse_option_menu = tk.OptionMenu(mouse_frame, self.mouse_action_type_var, *mouse_options)
        self.mouse_option_menu.config(bg=theme['MAIN_BG'], fg=theme['ENTRY_FG'], activebackground=theme['BTN_BG_HOVER'], relief=tk.FLAT, highlightthickness=0)
        self.mouse_option_menu["menu"].config(bg=theme['MAIN_BG'], fg=theme['ENTRY_FG'])
        self.mouse_option_menu.pack(side=tk.LEFT, padx=5)
        key_frame = tk.Frame(action_frame, bg=theme['MAIN_BG'])
        key_frame.pack(fill='x', pady=2)
        tk.Checkbutton(key_frame, text='键盘输入  ', variable=self.action_key_var,
            bg=theme['MAIN_BG'], fg=theme['MAIN_FG'],
            selectcolor=theme['MAIN_BG'],
            activebackground=theme['MAIN_BG'], activeforeground=theme['MAIN_FG'],
            highlightthickness=0, bd=0,
            command=self.toggle_key_input).pack(side=tk.LEFT)
        self.key_entry = tk.Entry(key_frame, font=('Segoe UI', 10), bg=theme['MAIN_BG'], fg=theme['ENTRY_FG'], relief=tk.FLAT, width=15, state=tk.DISABLED)
        self.key_entry.pack(side=tk.LEFT, padx=2)
        self.key_entry.insert(0, 'enter')
        sound_frame = tk.Frame(action_frame, bg=theme['MAIN_BG'])
        sound_frame.pack(fill='x', pady=2)
        tk.Checkbutton(sound_frame, text='播放提示音', variable=self.action_sound_var,
            bg=theme['MAIN_BG'], fg=theme['MAIN_FG'],
            selectcolor=theme['MAIN_BG'],
            activebackground=theme['MAIN_BG'], activeforeground=theme['MAIN_FG'],
            highlightthickness=0, bd=0).pack(side=tk.LEFT)
        sound_styles = ['叮叮', '升调', '降调', '错误', '成功']
        self.sound_style_menu = tk.OptionMenu(sound_frame, self.sound_style_var, *sound_styles)
        self.sound_style_menu.config(bg=theme['MAIN_BG'], fg=theme['ENTRY_FG'], activebackground=theme['BTN_BG_HOVER'], relief=tk.FLAT, highlightthickness=0, width=8)
        self.sound_style_menu["menu"].config(bg=theme['MAIN_BG'], fg=theme['ENTRY_FG'])
        self.sound_style_menu.pack(side=tk.LEFT, padx=5)
        # --- 2. 任务设置 ---
        options_frame = tk.LabelFrame(parent, text='任务与快捷键', bg=theme['MAIN_BG'], fg=theme['MAIN_FG'], font=('Segoe UI', 12, 'bold'), padx=10, pady=10, bd=1, relief=tk.GROOVE)
        options_frame.pack(pady=15, padx=20, fill='x')
        stop_on_find_frame = tk.Frame(options_frame, bg=theme['MAIN_BG'])
        stop_on_find_frame.pack(anchor='w', fill='x')
        tk.Checkbutton(stop_on_find_frame, text='找到目标后停止任务', variable=self.stop_on_find_var,
            bg=theme['MAIN_BG'], fg=theme['MAIN_FG'],
            selectcolor=theme['MAIN_BG'],
            activebackground=theme['MAIN_BG'], activeforeground=theme['MAIN_FG'],
            highlightthickness=0, bd=0).pack(anchor='w')
        confidence_frame = tk.Frame(options_frame, bg=theme['MAIN_BG'])
        confidence_frame.pack(anchor='w', pady=5, fill='x')
        tk.Label(confidence_frame, text="图像相似度:", bg=theme['MAIN_BG'], fg=theme['MAIN_FG']).pack(side=tk.LEFT, padx=(0, 5))
        confidence_slider = tk.Scale(confidence_frame, from_=0.1, to=1.0, resolution=0.01, orient=tk.HORIZONTAL, variable=self.confidence_var, bg=theme['MAIN_BG'], fg=theme['MAIN_FG'], troughcolor=theme['TAB_BG'], highlightthickness=0, length=250)
        confidence_slider.pack(side=tk.LEFT)
        interval_frame = tk.Frame(options_frame, bg=theme['MAIN_BG'])
        interval_frame.pack(anchor='w', pady=5, fill='x')
        tk.Label(interval_frame, text="检测间隔 (秒):", bg=theme['MAIN_BG'], fg=theme['MAIN_FG']).pack(side=tk.LEFT, padx=(0, 5))
        interval_entry = tk.Entry(interval_frame, textvariable=self.interval_var, font=('Segoe UI', 10), bg=theme['MAIN_BG'], fg=theme['ENTRY_FG'], relief=tk.FLAT, width=10)
        interval_entry.pack(side=tk.LEFT)
        interval_entry.config(bg=theme['MAIN_BG'])
        hotkey_frame = tk.Frame(options_frame, bg=theme['MAIN_BG'])
        hotkey_frame.pack(anchor='w', pady=5)
        tk.Label(hotkey_frame, text="启动快捷键:", bg=theme['MAIN_BG'], fg=theme['MAIN_FG']).pack(side=tk.LEFT, padx=(0, 5))
        start_entry = tk.Entry(hotkey_frame, textvariable=self.start_hotkey_var, font=('Segoe UI', 10), bg=theme['ENTRY_BG'], fg=theme['ENTRY_FG'], relief=tk.FLAT, width=10)
        start_entry.pack(side=tk.LEFT)
        start_entry.config(bg=theme['MAIN_BG'])
        tk.Label(hotkey_frame, text="停止快捷键:", bg=theme['MAIN_BG'], fg=theme['MAIN_FG']).pack(side=tk.LEFT, padx=(15, 5))
        stop_entry = tk.Entry(hotkey_frame, textvariable=self.stop_hotkey_var, font=('Segoe UI', 10), bg=theme['ENTRY_BG'], fg=theme['ENTRY_FG'], relief=tk.FLAT, width=10)
        stop_entry.pack(side=tk.LEFT)
        stop_entry.config(bg=theme['MAIN_BG'])
        ttk.Button(hotkey_frame, text="应用", command=self.apply_hotkeys, style='Rounded.TButton').pack(side=tk.LEFT, padx=10)
        
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

    def remove_target_by_index(self, idx):
        if 0 <= idx < len(self.targets):
            self.targets.pop(idx)
            self.log(f"已删除目标 {idx + 1}。")
            self.update_target_list_ui()
        else:
            self.log("请先从列表中选择一个要删除的目标。")

    def update_target_list_ui(self):
        theme = self._theme
        # Clear existing widgets
        for widget in self.target_list_frame.winfo_children():
            widget.destroy()
        if not self.targets:
            no_img_label = tk.Label(self.target_list_frame, text="请通过下方按钮添加目标图像", bg=theme['MAIN_BG'], fg='#888888', font=('Segoe UI', 10))
            no_img_label.pack(pady=20)
            return
        for i, target in enumerate(self.targets):
            item_frame = tk.Frame(self.target_list_frame, bg=theme['MAIN_BG'], bd=1, relief=tk.SOLID, highlightbackground=theme['BORDER_COLOR'], highlightcolor=theme['BORDER_COLOR'], highlightthickness=1)
            item_frame.pack(fill='x', expand=True, padx=5, pady=3)
            # --- Multi-select Checkbutton ---
            if 'selected_var' not in target:
                target['selected_var'] = tk.BooleanVar(value=target.get('selected', True))
            cb = tk.Checkbutton(
                item_frame,
                variable=target['selected_var'],
                bg=theme['MAIN_BG'],
                activebackground=theme['MAIN_BG'],
                fg=theme['MAIN_FG'],
                activeforeground=theme['MAIN_FG'],
                selectcolor=theme['MAIN_BG'],  # 选中时也是白色
                highlightthickness=0,
                bd=0,
                command=self.save_selected_state
            )
            cb.grid(row=0, column=0, padx=(5, 2), pady=5, sticky='w')
            # --- Thumbnail ---
            self.update_target_thumbnail(target)
            thumb_frame = tk.Frame(item_frame, width=120, height=60, bg=theme['MAIN_BG'])
            thumb_frame.grid_propagate(False)
            thumb_frame.grid(row=0, column=1, padx=(2, 8), pady=5, sticky='w')
            preview_label = tk.Label(thumb_frame, image=target['image_tk'], bg=theme['MAIN_BG'], anchor='center')
            preview_label.place(relx=0.5, rely=0.5, anchor='center')
            preview_label.bind("<Button-1>", lambda e, index=i: self.on_preview_click(e, index))
            # --- Offset Info ---
            offset_text = f"鼠标点击位置: {target['offset']}" if target['offset'] else "鼠标点击位置: 中心"
            offset_label = tk.Label(item_frame, text=offset_text, bg=theme['MAIN_BG'], fg=theme['MAIN_FG'])
            offset_label.grid(row=0, column=2, padx=(0, 8), pady=5, sticky='w')
            # --- Delete Button ---
            del_btn = ttk.Button(item_frame, text='删除', command=lambda idx=i: self.remove_target_by_index(idx), style='Danger.TButton')
            del_btn.grid(row=0, column=3, padx=(0, 10), pady=5, sticky='e')
            item_frame.grid_columnconfigure(0, weight=0)
            item_frame.grid_columnconfigure(1, weight=0)
            item_frame.grid_columnconfigure(2, weight=1)
            item_frame.grid_columnconfigure(3, weight=0)
        
    def save_selected_state(self):
        # 只更新内存，不立即保存到文件，等save_config时统一保存
        for target in self.targets:
            target['selected'] = target['selected_var'].get()

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
        self.log(f"目标 {target_index + 1} 的鼠标点击位置已更新为: {target['offset']}")
        self.update_target_list_ui() # Redraw to show the new crosshair

    def add_target_image(self, image_path_or_data):
        try:
            image = Image.open(image_path_or_data) if isinstance(image_path_or_data, str) else image_path_or_data
            new_target = {
                'image': image,
                'image_tk': None,
                'offset': None,
                'selected': True,  # 新增，默认选中
            }
            self.targets.append(new_target)

            self.log(f"已添加新目标，当前共 {len(self.targets)} 个目标。")
            self.update_target_list_ui()
        except Exception as e:
            self.log(f"错误: 无法加载图像 - {e}")

    def update_target_thumbnail(self, target):
        """Updates a target's thumbnail image including the crosshair."""
        MAX_THUMB_SIZE = (120, 60)  # 限制最大宽度和高度，防止长图撑爆
        thumb = target['image'].copy()
        thumb.thumbnail(MAX_THUMB_SIZE, Image.LANCZOS)
        
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
            # 只搜索选中的目标
            if target.get('selected', True):
                active_targets_info.append({
                    'index': i,
                    'image': target['image'],
                    'offset': target['offset']
                })
        
        if not active_targets_info:
            self.log("错误: 请先勾选至少一个目标图像。")
            return

        # Hide UI list by replacing its content
        for widget in self.target_list_frame.winfo_children():
            widget.destroy()
        hidden_label = tk.Label(self.target_list_frame, text="\n...任务运行中...\n", bg='#3C3C3C', fg='grey', font=('Arial', 14))
        hidden_label.pack(fill='x', pady=50)

        action_mouse = self.action_mouse_var.get()
        mouse_action_type = self.mouse_action_type_var.get()
        action_key = self.action_key_var.get()
        action_sound = self.action_sound_var.get()
        key_to_press = self.key_entry.get()
        stop_on_find = self.stop_on_find_var.get()
        confidence_level = self.confidence_var.get()
        interval = self.interval_var.get()

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        self.stop_event.clear()
        self.monitor_thread = threading.Thread(
            target=self.monitor_loop,
            args=(action_mouse, mouse_action_type, action_key, key_to_press, action_sound, stop_on_find, confidence_level, interval, active_targets_info),
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
        self.update_target_list_ui()
        
        self.log("任务已停止。")

    def monitor_loop(self, action_mouse, mouse_action_type, action_key, key_to_press, action_sound, stop_on_find, confidence_level, interval, active_targets):
        self.log(f"任务已启动，正在扫描屏幕 (相似度阈值: {confidence_level:.2f}, 间隔: {interval}s)...")
        while not self.stop_event.is_set():
            try:
                found_match_in_scan = False
                for target_info in active_targets:
                    if self.stop_event.is_set():
                        break

                    location_box = self.find_image_on_screen(target_info['image'], confidence=confidence_level)
                    
                    if location_box:
                        if self.stop_event.is_set():
                            break

                        if target_info['offset']:
                            target_x = location_box.left + target_info['offset'][0]
                            target_y = location_box.top + target_info['offset'][1]
                        else: # Default to center
                            target_x = location_box.left + location_box.width / 2
                            target_y = location_box.top + location_box.height / 2
                        
                        target_point = (int(target_x), int(target_y))
                        self.log(f"已找到目标 {target_info['index'] + 1}，鼠标点击位置: {target_point}")

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
                        
                        if action_sound:
                            try:
                                style = self.sound_style_var.get()
                                if style == '叮叮':
                                    sound_dingding()
                                elif style == '升调':
                                    sound_up()
                                elif style == '降调':
                                    sound_down()
                                elif style == '错误':
                                    sound_error()
                                elif style == '成功':
                                    sound_success()
                                elif style in self.sound_data:
                                    playsound(self.sound_data[style], block=False)
                                self.log(f"已播放声音：{style}")
                            except Exception as e:
                                tb_str = traceback.format_exc()
                                self.log(f"严重错误: 系统Beep播放失败. 错误: {e}\nTraceback:\n{tb_str}")

                        if action_key:
                            pyautogui.press(key_to_press)
                            self.log(f"已执行: 键盘输入 '{key_to_press}'。")
                        
                        found_match_in_scan = True
                        break # Stop scanning other targets in this cycle
                
                if self.stop_event.is_set():
                    break

                if found_match_in_scan:
                    if stop_on_find:
                        self.log("操作完成，任务即将停止。")
                        self.log_queue.put("TASK_COMPLETE")
                        break
                    else:
                        self.log("操作完成，继续扫描...")
                        time.sleep(interval)
                else:
                    # To avoid log spam, we check the stop event before logging the "not found" message and sleeping.
                    if not self.stop_event.is_set():
                        self.log(f"未在屏幕上找到任何目标，{interval}秒后重试...")
                        time.sleep(interval)

            except Exception as e:
                self.log(f"监控循环发生严重错误: {e}")
                if not self.stop_event.is_set():
                    time.sleep(interval) # Wait before retrying

    def load_available_sounds(self):
        """扫描并加载可用的声音文件。"""
        # 始终包含内置声音选项
        self.available_sounds = ["叮叮", "升调", "降调", "错误", "成功", "无"]
        self.sound_data = {} # Clear previous data
        
        # 扫描并加载本地.wav文件
        try:
            scan_dir = '.'
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                scan_dir = sys._MEIPASS

            local_wavs = [f for f in os.listdir(scan_dir) if f.lower().endswith('.wav')]
            if local_wavs:
                for wav_file in local_wavs:
                    try:
                        # Extract the sound from the package to a real temporary file
                        packaged_path = resource_path(wav_file)
                        with open(packaged_path, 'rb') as f_in:
                            data = f_in.read()
                        
                        with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as temp_f_out:
                            temp_f_out.write(data)
                            real_path = temp_f_out.name

                        self.sound_data[wav_file] = real_path
                        self.temp_sound_files.append(real_path)
                        # wav文件追加到菜单末尾
                        self.available_sounds.append(wav_file)
                        self.log(f"信息: 已解包 '{wav_file}' 到临时文件 '{os.path.basename(real_path)}'。")

                    except Exception as e:
                        self.log(f"警告: 无法处理WAV文件 '{wav_file}' - {e}")
        except Exception as e:
            self.log(f"警告: 扫描本地WAV文件失败 - {e}")
        
        # 更新UI
        if hasattr(self, 'sound_style_menu'):
            menu = self.sound_style_menu['menu']
            menu.delete(0, 'end')
            for sound in self.available_sounds:
                menu.add_command(label=sound, command=lambda value=sound: self.sound_style_var.set(value))
            
            current_sound = self.sound_style_var.get()
            if current_sound not in self.available_sounds:
                self.sound_style_var.set(self.available_sounds[0])
        else:
            self.log("错误: sound_style_menu 未初始化。")

    def toggle_mouse_input(self):
        state = tk.NORMAL if self.action_mouse_var.get() else tk.DISABLED
        self.mouse_option_menu.config(state=state)

    def toggle_key_input(self):
        if self.action_key_var.get():
            self.key_entry.config(state=tk.NORMAL)
        else:
            self.key_entry.config(state=tk.DISABLED)

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

    def save_config(self):
        """Saves current settings and targets to a JSON file."""
        if not os.path.exists(self.targets_dir):
            os.makedirs(self.targets_dir)

        target_configs = []
        for i, target in enumerate(self.targets):
            image_path = os.path.join(self.targets_dir, f"target_{i}.png")
            target['image'].save(image_path)
            target_configs.append({
                'image_path': image_path,
                'offset': target['offset'],
                'selected': target.get('selected', True)  # 保存选中状态
            })

        config = {
            "action_mouse": self.action_mouse_var.get(),
            "mouse_action_type": self.mouse_action_type_var.get(),
            "action_key": self.action_key_var.get(),
            "key_entry": self.key_entry.get(),
            "action_sound": self.action_sound_var.get(),
            "sound_style": self.sound_style_var.get(),
            "stop_on_find": self.stop_on_find_var.get(),
            "confidence": self.confidence_var.get(),
            "interval": self.interval_var.get(),
            "start_hotkey": self.start_hotkey_var.get(),
            "stop_hotkey": self.stop_hotkey_var.get(),
            "targets": target_configs
        }

        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)
            self.log("配置已成功保存。")
        except Exception as e:
            self.log(f"错误: 保存配置失败 - {e}")

    def load_config(self):
        """Loads settings and targets from the JSON file."""
        if not os.path.exists(self.config_file):
            self.log("未找到配置文件，使用默认设置。")
            return
        
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)

            self.action_mouse_var.set(config.get("action_mouse", True))
            self.mouse_action_type_var.set(config.get("mouse_action_type", "左键单击"))
            self.action_key_var.set(config.get("action_key", False))
            self.key_entry.delete(0, tk.END)
            self.key_entry.insert(0, config.get("key_entry", "enter"))
            self.action_sound_var.set(config.get("action_sound", True))
            self.sound_style_var.set(config.get("sound_style", "叮叮"))
            self.stop_on_find_var.set(config.get("stop_on_find", False))
            self.confidence_var.set(config.get("confidence", 0.9))
            self.interval_var.set(config.get("interval", 5.0))
            self.start_hotkey_var.set(config.get("start_hotkey", "Alt-1"))
            self.stop_hotkey_var.set(config.get("stop_hotkey", "Alt-2"))

            # Manually update UI state based on loaded values
            self.toggle_mouse_input()
            self.toggle_key_input()

            for target_config in config.get("targets", []):
                image_path = target_config.get("image_path")
                if image_path and os.path.exists(image_path):
                    new_target = {
                        'image': Image.open(image_path),
                        'image_tk': None,
                        'offset': target_config.get('offset'),
                        'selected': target_config.get('selected', True),  # 恢复选中状态
                    }
                    self.targets.append(new_target)
            
            self.update_target_list_ui()
            self.log("配置已成功加载。")

        except Exception as e:
            self.log(f"错误: 加载配置失败 - {e}")

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
        self.canvas.bind("<Button-3>", self.on_right_click_cancel)  # 新增右键取消
        self.canvas.bind("<ButtonRelease-3>", self.on_right_click_cancel)  # 防止事件穿透

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
        if (x2 - x1) > 0 and (y2 - y1) > 0:
            self.region = (int(x1), int(y1), int(x2), int(y2))
        else:
            self.region = None # Invalid region
        self.root.quit()
        self.root.destroy()

    def on_right_click_cancel(self, event):

        self.region = None
        self.root.quit()
        self.root.destroy()
        return "break"

    def get_region(self):
        self.root.mainloop()
        return self.region

def main():
    root = tk.Tk()
    app = ImageAutoClickerApp(root)
    root.mainloop()

if __name__ == '__main__':
    main() 