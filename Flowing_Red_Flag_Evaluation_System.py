# 流动红旗评比系统 (Flowing Red Flag Evaluation System)
# Copyright (C) 2025 流动红旗评比系统开发团队
#
# 本程序是自由软件：您可以根据自由软件基金会发布的GNU Affero通用公共许可证版本3
# 或（您选择的）任何 later version 的条款重新分发和/或修改它。
#
# 本程序的分发是希望它能发挥作用，但不提供任何担保；甚至没有适销性或特定用途适用性的隐含保证。
# 有关详细信息，请参阅GNU Affero通用公共许可证。
#
# 您应该已经收到了GNU Affero通用公共许可证的副本。如果没有，请参见<http://www.gnu.org/licenses/>。
#
# 本项目采用 Creative Commons Attribution-ShareAlike 3.0 Unported License (CC-BY-SA 3.0)
# 您可以自由地共享和演绎本作品，但必须署名并以相同方式共享。
#
# 更多信息请查看项目根目录的 LICENSE 文件和 README.md 文件。

import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog
import json
import datetime
import os

def setup_chinese_font_support():
    try:
        import tkinter.font as tkFont
        root = tk.Tk()
        available_fonts = list(tkFont.families())
        root.destroy()
        preferred_fonts = [
            "Microsoft YaHei UI", "Microsoft YaHei", "SimHei", "KaiTi", 
            "FangSong", "STHeiti", "STSong", "PingFang SC", "Hiragino Sans GB",
            "Noto Sans CJK SC", "Source Han Sans CN", "Arial Unicode MS"
        ]
        for font in preferred_fonts:
            if font in available_fonts:
                return font
        return "Arial Unicode MS" if "Arial Unicode MS" in available_fonts else None
    except Exception as e:
        print(f"字体设置出错: {e}")
        return None

chinese_font = setup_chinese_font_support()

class Config:
    ITEMS = {
        "早迟到": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 10},
        "早读": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 10},
        "节能开窗": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 5},
        "仪容仪表": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 5},
        "跑操": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 30},
        "午休": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 10},
        "卫生": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 30},
        "巡视": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 10},
        "及时上交文件": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 5},
        "宿舍": {"columns": ("班级", "周一", "周二", "周三", "周四", "周五", "平均分"), "max_score": 5}
    }
    
    DUAL_PERIOD_ITEMS = {"跑操", "卫生"}
    # Default classes - will be overridden by settings
    CLASSES = [f"高二{i}班" for i in range(1, 11)]
    # Default weighted addition - will be overridden by settings
    WEIGHTED_ADDITION = {
        "高二1班": 0, "高二2班": 0.5, **{f"高二{i}班": 2 for i in range(3, 11)}
    }

class SettingsManager:
    def __init__(self):
        self.settings_file = "settings.json"
        self.settings = {
            "root_directory": os.getcwd(),
            "max_scores": {name: info["max_score"] for name, info in Config.ITEMS.items()},
            "classes": Config.CLASSES,
            "weighted_addition": Config.WEIGHTED_ADDITION
        }
        self.load_settings()
    
    def load_settings(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    self.settings.update(loaded_settings)
                    
                    # Ensure classes and weighted_addition exist in settings
                    if "classes" not in self.settings:
                        self.settings["classes"] = Config.CLASSES
                    if "weighted_addition" not in self.settings:
                        self.settings["weighted_addition"] = Config.WEIGHTED_ADDITION
        except Exception as e:
            print(f"加载设置时出错: {e}")
    
    def save_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存设置时出错: {e}")
    
    def set_root_directory(self, directory):
        self.settings["root_directory"] = directory
        self.save_settings()
    
    def get_root_directory(self):
        return self.settings.get("root_directory", os.getcwd())
    
    def get_data_file_path(self):
        """获取数据文件目录路径"""
        return self.get_root_directory()
    
    def get_log_file_path(self):
        return os.path.join(self.get_root_directory(), "flowing_flag_log.txt")
    
    def get_max_score(self, item_name):
        return self.settings["max_scores"].get(item_name, Config.ITEMS[item_name]["max_score"])
    
    def set_max_score(self, item_name, max_score):
        self.settings["max_scores"][item_name] = max_score
        self.save_settings()
    
    def get_classes(self):
        return self.settings.get("classes", Config.CLASSES)
    
    def set_classes(self, classes):
        self.settings["classes"] = classes
        self.save_settings()
    
    def get_weighted_addition(self):
        return self.settings.get("weighted_addition", Config.WEIGHTED_ADDITION)
    
    def set_weighted_addition(self, weighted_addition):
        self.settings["weighted_addition"] = weighted_addition
        self.save_settings()

class LogManager:
    def __init__(self, settings_manager):
        self.settings_manager = settings_manager
        self.ensure_log_directory()
    
    def ensure_log_directory(self):
        root_dir = self.settings_manager.get_root_directory()
        if not os.path.exists(root_dir):
            try:
                os.makedirs(root_dir)
            except Exception as e:
                messagebox.showerror("错误", f"创建目录失败: {e}")
    
    def log(self, message):
        try:
            log_file = self.settings_manager.get_log_file_path()
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = f"[{timestamp}] {message}\n"
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(log_entry)
        except Exception as e:
            print(f"记录日志时出错: {e}")

class HistoryManager:
    def __init__(self):
        self.history = []
        self.current_index = -1
        self.max_history = 50
    
    def add_record(self, data):
        """添加历史记录"""
        self.history = self.history[:self.current_index + 1]
        date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.history.append({'date': date, 'data': data})
        self.current_index = len(self.history) - 1
        
        if len(self.history) > self.max_history:
            self.history.pop(0)
            self.current_index -= 1
    
    def get_history(self):
        return self.history
    
    def can_undo(self):
        """检查是否可以撤销"""
        return self.current_index > 0
    
    def can_redo(self):
        """检查是否可以重做"""
        return self.current_index < len(self.history) - 1
    
    def undo(self):
        """撤销操作"""
        if self.can_undo():
            self.current_index -= 1
            return self.history[self.current_index]['data']
        return None
    
    def redo(self):
        """重做操作"""
        if self.can_redo():
            self.current_index += 1
            return self.history[self.current_index]['data']
        return None
    
    def get_current_data(self):
        """获取当前数据"""
        if self.current_index >= 0 and self.current_index < len(self.history):
            return self.history[self.current_index]['data']
        return None
    
    def clear_history(self):
        """清空历史记录"""
        self.history = []
        self.current_index = -1

class FlowingRedFlagEvaluationSystem:
    def __init__(self, root):
        self.root = root
        self.settings_manager = SettingsManager()
        self.root.title('流动红旗评比系统')
        self.root.geometry("1600x900")
        
        self.log_manager = LogManager(self.settings_manager)
        self.history_manager = HistoryManager()
        
        self.punishments = {}
        self.items = Config.ITEMS
        self.classes = self.settings_manager.get_classes()
        self.weighted_addition = self.settings_manager.get_weighted_addition().copy()
        self.dual_period_items = Config.DUAL_PERIOD_ITEMS
        
        # 初始化class_combobox为None
        self.class_combobox = None
        
        self.create_main_layout()
        self.create_menu()
        self.create_notebook()
        self.create_sidebar()
        self.create_status_bar()
        self.bind_shortcuts()
        
        self.punishment_list_tree = None
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)
        self.root.bind("<Button-1>", self.on_click_anywhere)
        
        self.save_snapshot()
        self.log_manager.log("系统启动")
        
        self.root.after(100, self.show_welcome_message)
    
    def show_welcome_message(self):
        self.update_status("🟢 系统已启动，双击表格单元格可编辑数据")
        
        # 检查是否是首次运行（没有保存的数据文件）
        if not self.has_data_file():
            self.prompt_for_data_file()
    
    def has_data_file(self):
        """检查是否存在数据文件"""
        root_dir = self.settings_manager.get_root_directory()
        # 检查目录下是否有JSON数据文件
        if os.path.exists(root_dir) and os.path.isdir(root_dir):
            for file in os.listdir(root_dir):
                if file.endswith('.json') and file.startswith('流动红旗分数数据_'):
                    return True
        return False
    
    def prompt_for_data_file(self):
        """提示用户选择数据文件"""
        result = messagebox.askyesno("首次运行", "检测到这是首次运行系统，是否要加载现有的数据文件？\n\n选择'是'加载数据文件，选择'否'使用默认设置。")
        if result:
            self.load_data()
    
    def create_main_layout(self):
        style = ttk.Style()
        print("可用的主题:", style.theme_names())
        
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.title_frame = ttk.Frame(self.main_frame)
        self.title_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = ttk.Label(self.title_frame, text='流动红旗评比系统', 
                               font=("微软雅黑", 18, "bold"))
        title_label.pack()
        
        # 添加分割线
        separator = ttk.Separator(self.main_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 10))
        
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单 - 数据操作相关
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='文件', menu=file_menu)
        file_menu.add_command(label='保存数据', command=self.save_data, accelerator="Ctrl+S")
        file_menu.add_command(label='另存为', command=self.save_as_data)
        file_menu.add_command(label='加载数据', command=self.load_data, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label='导出总分表格', command=self.save_total_score_table)
        file_menu.add_separator()
        file_menu.add_command(label='退出', command=self.root.quit)
        
        # 编辑菜单 - 数据编辑相关
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='编辑', menu=edit_menu)
        edit_menu.add_command(label='撤销', command=self.undo_action, accelerator="Ctrl+Z")
        edit_menu.add_command(label='重做', command=self.redo_action, accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label='复原数据', command=self.reset_program)
        
        # 视图菜单 - 界面视图相关
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='视图', menu=view_menu)
        view_menu.add_command(label='全屏切换', command=self.toggle_fullscreen, accelerator="F11")
        
        # 数据菜单 - 数据管理相关
        data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='数据', menu=data_menu)
        data_menu.add_command(label='奖惩管理', command=self.manage_punishments)
        data_menu.add_command(label='历史记录', command=self.show_history)
        
        # 设置菜单 - 系统设置相关
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='设置', menu=settings_menu)
        settings_menu.add_command(label='系统设置', command=self.open_settings)
        
        # 工具菜单 - 计算和输出相关
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='工具', menu=tools_menu)
        tools_menu.add_command(label='计算总分', command=self.calculate_totals)
        tools_menu.add_command(label='输出表格', command=self.calculate_and_output_table)
        tools_menu.add_command(label='评比结果', command=self.show_evaluation_result)
        tools_menu.add_command(label='导出表格', command=self.save_total_score_table)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='帮助', menu=help_menu)
        help_menu.add_command(label='关于系统', command=self.show_about)
    
    def create_notebook(self):
        self.left_frame = ttk.Frame(self.content_frame)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.notebook = ttk.Notebook(self.left_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        self.pages = {}
        for name in self.items:
            self.pages[name] = self.create_page(name)
        
        for name, page in self.pages.items():
            self.notebook.add(page, text=name)
    
    def create_page(self, page_name):
        frame = ttk.Frame(self.notebook)
        frame.pack(fill=tk.BOTH, expand=True)
        
        item_info = self.items[page_name]
        columns = item_info["columns"]
        max_score = self.settings_manager.get_max_score(page_name)
        
        if page_name in self.dual_period_items:
            am_frame = ttk.LabelFrame(frame, text=f"上午{page_name}")
            am_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            am_tree = self.create_tree(am_frame, columns)
            am_tree.pack(fill=tk.BOTH, expand=True)
            
            pm_frame = ttk.LabelFrame(frame, text=f"下午{page_name}")
            pm_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            pm_tree = self.create_tree(pm_frame, columns)
            pm_tree.pack(fill=tk.BOTH, expand=True)
            
            frame.am_tree = am_tree
            frame.pm_tree = pm_tree
            
            for cls in self.classes:
                values = (cls, max_score, max_score, max_score, max_score, max_score, max_score)
                am_tree.insert("", "end", values=values)
                pm_tree.insert("", "end", values=values)
        else:
            tree = self.create_tree(frame, columns)
            tree.pack(fill=tk.BOTH, expand=True)
            
            for cls in self.classes:
                values = (cls, max_score, max_score, max_score, max_score, max_score, max_score)
                tree.insert("", "end", values=values)
        
        return frame
    
    def create_tree(self, parent, columns):
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=80, anchor="center")
        
        tree.bind("<Double-1>", lambda event, tree=tree: self.on_double_click(event, tree))
        return tree
    
    def create_sidebar(self):
        self.right_frame = ttk.Frame(self.content_frame)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        
        self.total_frame = ttk.LabelFrame(self.right_frame, text="🏆 总分排名", padding=10)
        self.total_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        tree_container = ttk.Frame(self.total_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        self.total_tree = ttk.Treeview(tree_container, columns=("排名", "班级", "总分"), show="headings", height=12)
        self.total_tree.heading("排名", text="排名")
        self.total_tree.heading("班级", text="班级")
        self.total_tree.heading("总分", text="总分")
        self.total_tree.column("排名", width=40, anchor="center")
        self.total_tree.column("班级", width=90, anchor="center")
        self.total_tree.column("总分", width=70, anchor="center")
        
        tree_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.total_tree.yview)
        self.total_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        self.total_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.total_tree.tag_configure("first_place", background="#FFD700", foreground="black")
        self.total_tree.tag_configure("top_five", background="#87CEEB", foreground="black")
        self.total_tree.tag_configure("normal", background="white", foreground="black")
        
        self.action_frame = ttk.LabelFrame(self.right_frame, text="🛠️ 操作面板", padding=10)
        self.action_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 合并计算总分和刷新功能的按钮
        calculate_refresh_buttons = [
            ("📊 计算/刷新", self.calculate_totals),
            ("📋 输出表格", self.calculate_and_output_table)
        ]
        self.create_button_row(self.action_frame, calculate_refresh_buttons)
        
        # 奖惩管理和复原数据按钮
        management_buttons = [
            ("⚖️ 奖惩管理", self.manage_punishments),
            ("🔄 复原数据", self.reset_program)
        ]
        self.create_button_row(self.action_frame, management_buttons)
        
        # 评比结果和导出表格按钮
        result_buttons = [
            ("🏆 评比结果", self.show_evaluation_result),
            ("💾 导出表格", self.save_total_score_table)
        ]
        self.create_button_row(self.action_frame, result_buttons)
    
    def create_button_row(self, parent, buttons):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, padx=5, pady=5)
        
        for text, command in buttons:
            btn = ttk.Button(frame, text=text, command=command)
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=3, pady=2)
        return frame
    
    def create_status_bar(self):
        self.status_frame = ttk.Frame(self.root, relief=tk.RAISED, borderwidth=1)
        self.status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
        
        self.status_icon = ttk.Label(self.status_frame, text="🟢", anchor=tk.W, width=3, font=("微软雅黑", 10), background="#e8f4f8")
        self.status_icon.pack(side=tk.LEFT, padx=(5, 10))
        
        self.status_bar = ttk.Label(self.status_frame, text="系统就绪", relief=tk.FLAT, anchor=tk.W, 
                                   font=("微软雅黑", 9), foreground="blue", background="#e8f4f8")
        self.status_bar.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 10))
        
        self.tip_label = ttk.Label(self.status_frame, text="提示: 双击表格可编辑数据 | Ctrl+T: 输出表格 | Ctrl+P: 惩罚管理 | Ctrl+R: 复原数据", 
                                  relief=tk.FLAT, anchor=tk.CENTER, foreground="gray", 
                                  font=("微软雅黑", 8), background="#e8f4f8")
        self.tip_label.pack(side=tk.LEFT, padx=(10, 10))
        
        self.time_label = ttk.Label(self.status_frame, text="", relief=tk.FLAT, anchor=tk.E, 
                                   font=("微软雅黑", 9), foreground="darkgreen", background="#e8f4f8")
        self.time_label.pack(fill=tk.X, side=tk.RIGHT, expand=True, padx=(10, 5))
        self.update_time()
    
    def update_time(self):
        if self.root.winfo_exists():
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.time_label.config(text=current_time)
            self.time_after_id = self.root.after(1000, self.update_time)
    
    def on_double_click(self, event, tree):
        region = tree.identify_region(event.x, event.y)
        if region == "cell":
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            
            if column == "#1":
                return
            current_value = tree.item(item, "values")[int(column[1]) - 1]
            
            bbox = tree.bbox(item, column)
            if not bbox:
                return
            
            x, y, width, height = bbox
            
            entry = ttk.Entry(tree, justify='center', font=("微软雅黑", 9))
            entry.place(x=x, y=y, width=width, height=height)
            entry.insert(0, current_value)
            entry.select_range(0, tk.END)
            
            entry.bind("<Return>", lambda e, item=item, column=column, tree=tree, entry=entry: 
                       self.on_enter(e, item, column, tree, entry))
            entry.bind("<Escape>", lambda e, entry=entry: self.on_escape(e, entry))
            entry.bind("<FocusOut>", lambda e, item=item, column=column, tree=tree, entry=entry: 
                       self.on_enter(e, item, column, tree, entry))
            entry.focus_set()
            
            try:
                values = list(tree.item(item, "values"))
                self.update_status(f"正在编辑: {values[0]} 班级 {column} 项目")
            except Exception as e:
                self.update_status(f"编辑项目时出错: {str(e)}")
    
    def on_enter(self, event, item, column, tree, entry):
        new_value = entry.get()
        try:
            page_name = self.notebook.tab(self.notebook.select(), "text")
            
            for widget in self.notebook.nametowidget(self.notebook.select()).winfo_children():
                if hasattr(widget, 'am_tree') and widget.am_tree == tree:
                    page_name = widget.cget('text').replace('上午', '')
                    break
                elif hasattr(widget, 'pm_tree') and widget.pm_tree == tree:
                    page_name = widget.cget('text').replace('下午', '')
                    break
            
            score = float(new_value)
            max_score = self.settings_manager.get_max_score(page_name)
            if score < 0 or score > max_score:
                messagebox.showerror("输入错误", f"分数应在0-{max_score}之间！\n\n您输入的值: {new_value}")
                self.update_status(f"输入错误: 分数应在0-{max_score}之间")
                return
        except ValueError:
            messagebox.showerror("输入错误", f"请输入有效的数字！\n\n您输入的值: {new_value}")
            self.update_status("输入错误: 请输入有效的数字")
            return
        
        try:
            values = list(tree.item(item, "values"))
            col_index = int(column[1]) - 1
            values[col_index] = new_value
            
            if 1 <= col_index <= 5:
                scores = list(map(float, values[1:6]))
                avg_score = sum(scores) / 5
                values[6] = round(avg_score, 2)
            
            tree.item(item, values=values)
            
            self.save_snapshot()
            
            entry.destroy()
            self.update_status(f"已更新 {values[0]} 班级的分数")
            self.log_manager.log(f"更新 {values[0]} 班级 {page_name} 项目分数为 {new_value}")
            self.calculate_totals()
        except Exception as e:
            self.update_status(f"更新分数时出错: {str(e)}")
            entry.destroy()
    
    def on_escape(self, event, entry):
        entry.destroy()
    
    def update_status(self, message):
        self.status_bar.config(text=message)
        self.root.after(3000, lambda: self.status_bar.config(text="就绪"))
    
    def bind_shortcuts(self):
        self.root.bind('<Control-s>', lambda e: self.save_data())
        self.root.bind('<Control-o>', lambda e: self.load_data())
        self.root.bind('<Control-t>', lambda e: self.calculate_and_output_table())
        self.root.bind('<Control-p>', lambda e: self.manage_punishments())
        self.root.bind('<Control-r>', lambda e: self.reset_program())
        self.root.bind('<Control-e>', lambda e: self.save_total_score_table())
        self.root.bind('<Control-z>', lambda e: self.undo_action())
        self.root.bind('<Control-y>', lambda e: self.redo_action())
        self.root.bind('<F5>', lambda e: self.calculate_totals())
        self.root.bind('<F6>', lambda e: self.show_evaluation_result())
        self.root.bind('<F11>', lambda e: self.toggle_fullscreen())
    
    def reset_program(self):
        if messagebox.askyesno("确认", "确定要复原所有数据吗？"):
            self.save_snapshot()
            
            self.reset_data()
            self.punishments.clear()
            self.update_status("数据已复原")
            self.log_manager.log("执行数据复原操作")
    
    def reset_data(self):
        self.classes = self.settings_manager.get_classes()
        
        for page_name, page_frame in self.pages.items():
            if page_name in self.dual_period_items:
                self.reset_tree_data(page_frame.am_tree, page_name)
                self.reset_tree_data(page_frame.pm_tree, page_name)
            else:
                for widget in page_frame.winfo_children():
                    if isinstance(widget, ttk.Treeview):
                        self.reset_tree_data(widget, page_name)
        
        for item in self.total_tree.get_children():
            self.total_tree.delete(item)
        
        self.punishments.clear()
        # 检查punishment_list_tree组件是否仍然有效
        if hasattr(self, 'punishment_list_tree') and self.punishment_list_tree is not None:
            for item in self.punishment_list_tree.get_children():
                self.punishment_list_tree.delete(item)
    
    def reset_tree_data(self, tree, page_name):
        max_score = self.settings_manager.get_max_score(page_name)
        if isinstance(tree, ttk.Treeview):
            for item in tree.get_children():
                values = list(tree.item(item, "values"))
                for i in range(1, 6):
                    values[i] = max_score
                values[6] = max_score
                tree.item(item, values=values)
        else:
            for widget in tree.winfo_children():
                if isinstance(widget, ttk.Treeview):
                    for item in widget.get_children():
                        values = list(widget.item(item, "values"))
                        for i in range(1, 6):
                            values[i] = max_score
                        values[6] = max_score
                        widget.item(item, values=values)
    
    def calculate_totals(self):
        for item in self.total_tree.get_children():
            self.total_tree.delete(item)
        
        class_scores = {}
        
        for cls in self.classes:
            total_score = 0
            for page_name, page_frame in self.pages.items():
                max_score = self.settings_manager.get_max_score(page_name)
                
                if page_name in self.dual_period_items:
                    am_avg = self.get_class_average(page_frame.am_tree, cls)
                    pm_avg = self.get_class_average(page_frame.pm_tree, cls)
                    total_score += ((am_avg + pm_avg) / 2)
                else:
                    avg_score = self.get_class_average(page_frame, cls)
                    total_score += avg_score
            total_score += self.weighted_addition[cls]
            if cls in self.punishments:
                for punishment in self.punishments[cls]:
                    if punishment["type"] == "add":
                        total_score += punishment["score"]
                    elif punishment["type"] == "subtract":
                        total_score -= punishment["score"]
            
            class_scores[cls] = round(total_score, 2)
        
        sorted_classes = sorted(class_scores.items(), key=lambda x: x[1], reverse=True)
        
        for i, (cls, score) in enumerate(sorted_classes):
            rank = i + 1
            if i == 0:
                self.total_tree.insert("", "end", values=(rank, cls, score), tags=("first_place",))
            elif i < 5:
                self.total_tree.insert("", "end", values=(rank, cls, score), tags=("top_five",))
            else:
                self.total_tree.insert("", "end", values=(rank, cls, score))
        
        self.log_manager.log("执行总分计算")
    
    def get_class_average(self, tree, cls):
        if isinstance(tree, ttk.Treeview):
            for item in tree.get_children():
                values = tree.item(item, "values")
                if values[0] == cls:
                    return float(values[6])
        else:
            for widget in tree.winfo_children():
                if isinstance(widget, ttk.Treeview):
                    for item in widget.get_children():
                        values = widget.item(item, "values")
                        if values[0] == cls:
                            return float(values[6])
        return 0.0
    
    def calculate_and_output_table(self):
        self.calculate_totals()
        self.show_output_table()
    
    def show_output_table(self):
        output_window = tk.Toplevel(self.root)
        output_window.title("流动红旗评比总分表")
        output_window.geometry("1200x600")
        output_window.transient(self.root)
        
        columns = ("班级", "早迟到", "早读", "节能开窗", "仪容仪表", "间操", "午休", "上下午各班卫生", "巡视", "及时上交文件", "宿舍", "加权", "奖惩分", "奖惩备注", "总分")
        tree = ttk.Treeview(output_window, columns=columns, show="headings")
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=80)
        
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        for cls in self.classes:
            values = [cls]
            for page_name in self.items:
                max_score = self.settings_manager.get_max_score(page_name)
                
                if page_name in self.dual_period_items:
                    am_avg = self.get_class_average(self.pages[page_name].am_tree, cls)
                    pm_avg = self.get_class_average(self.pages[page_name].pm_tree, cls)
                    avg_score = (am_avg + pm_avg) / 2
                    values.append(round(avg_score, 2))
                else:
                    avg_score = self.get_class_average(self.pages[page_name], cls)
                    values.append(avg_score)
            values.append(self.weighted_addition[cls])
            punishment_score = 0
            punishment_notes = []
            if cls in self.punishments:
                for punishment in self.punishments[cls]:
                    if punishment["type"] == "add":
                        punishment_score += punishment["score"]
                    elif punishment["type"] == "subtract":
                        punishment_score -= punishment["score"]
                    punishment_notes.append(punishment["note"])
            values.append(punishment_score)
            values.append("\n".join(punishment_notes))
            total_score = sum(map(float, values[1:-2])) + values[-2]
            values.append(round(total_score, 2))
            
            tree.insert("", "end", values=values)
        
        v_scrollbar = ttk.Scrollbar(output_window, orient=tk.VERTICAL, command=tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(output_window, orient=tk.HORIZONTAL, command=tree.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        tree.configure(xscrollcommand=h_scrollbar.set)
    
    def show_evaluation_result(self):
        self.calculate_totals()
        
        class_scores = {}
        for item in self.total_tree.get_children():
            values = self.total_tree.item(item, "values")
            class_scores[values[1]] = float(values[2])
        
        sorted_classes = sorted(class_scores.items(), key=lambda x: x[1], reverse=True)
        
        result_window = tk.Toplevel(self.root)
        result_window.title("流动红旗评比结果")
        result_window.geometry("600x400")
        result_window.transient(self.root)
        
        main_frame = ttk.Frame(result_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        title_label = ttk.Label(main_frame, text="流动红旗评比结果", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(result_frame, text="校级流动红旗获得者：", font=("微软雅黑", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        ttk.Label(result_frame, text=f"{sorted_classes[0][0]}（总分：{sorted_classes[0][1]}）", font=("微软雅黑", 12)).pack(anchor=tk.W)
        
        ttk.Label(result_frame, text="年级流动红旗获得者：", font=("微软雅黑", 12, "bold")).pack(anchor=tk.W, pady=(20, 10))
        for i in range(1, min(5, len(sorted_classes))):
            ttk.Label(result_frame, text=f"{sorted_classes[i][0]}（总分：{sorted_classes[i][1]}）", font=("微软雅黑", 12)).pack(anchor=tk.W)
    
    def manage_punishments(self):
        # 使用局部变量而不是实例变量来避免组件引用问题
        local_vars = {}
        
        punishment_window = tk.Toplevel(self.root)
        punishment_window.title("奖惩管理")
        punishment_window.geometry("800x500")
        punishment_window.transient(self.root)
        punishment_window.grab_set()
        
        main_frame = ttk.Frame(punishment_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        input_frame = ttk.LabelFrame(main_frame, text="添加奖惩")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        input_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(input_frame, text="选择班级:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        # 使用局部变量存储组件引用
        local_vars['class_combobox'] = ttk.Combobox(input_frame, values=self.classes)
        local_vars['class_combobox'].grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        ttk.Label(input_frame, text="奖惩类型:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        local_vars['punishment_type'] = tk.StringVar()
        type_frame = ttk.Frame(input_frame)
        type_frame.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        ttk.Radiobutton(type_frame, text="奖励（加分）", variable=local_vars['punishment_type'], value="add").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="惩罚（减分）", variable=local_vars['punishment_type'], value="subtract").pack(side=tk.LEFT, padx=5)
        
        ttk.Label(input_frame, text="分值:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        local_vars['score_entry'] = ttk.Entry(input_frame)
        local_vars['score_entry'].grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        
        ttk.Label(input_frame, text="备注:").grid(row=3, column=0, padx=10, pady=10, sticky="e")
        local_vars['note_entry'] = ttk.Entry(input_frame)
        local_vars['note_entry'].grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        # 定义局部的添加和删除方法
        def add_punishment_local():
            cls = local_vars['class_combobox'].get()
            if not cls:
                messagebox.showerror("错误", "请选择班级！")
                return
            ptype = local_vars['punishment_type'].get()
            if not ptype:
                messagebox.showerror("错误", "请选择奖惩类型！")
                return
            score = local_vars['score_entry'].get()
            try:
                score = float(score)
            except ValueError:
                messagebox.showerror("错误", "请输入有效的分值！")
                return
            note = local_vars['note_entry'].get()
            if not note:
                messagebox.showerror("错误", "请输入备注！")
                return
            
            if cls not in self.punishments:
                self.punishments[cls] = []
            self.punishments[cls].append({"type": ptype, "score": score, "note": note})
            
            type_text = "奖励" if ptype == "add" else "惩罚"
            local_vars['punishment_list_tree'].insert("", "end", values=(cls, type_text, score, note))
            
            self.save_snapshot()
            
            local_vars['score_entry'].delete(0, tk.END)
            local_vars['note_entry'].delete(0, tk.END)
            self.update_status(f"已为 {cls} 添加奖惩")
            self.log_manager.log(f"为 {cls} 添加{ptype}分 {score}，备注: {note}")
        
        def delete_punishment_local():
            selected_item = local_vars['punishment_list_tree'].selection()
            if not selected_item:
                messagebox.showerror("错误", "请选择要删除的奖惩！")
                return
            
            self.save_snapshot()
            item = selected_item[0]
            values = local_vars['punishment_list_tree'].item(item, "values")
            cls = values[0]
            ptype = values[1]
            score = float(values[2])
            note = values[3]
            
            if cls in self.punishments:
                for i, punishment in enumerate(self.punishments[cls]):
                    if punishment["type"] == ptype and punishment["score"] == score and punishment["note"] == note:
                        self.punishments[cls].pop(i)
                        local_vars['punishment_list_tree'].delete(item)
                        
                        self.save_snapshot()
                        
                        self.update_status(f"已删除 {cls} 的奖惩记录")
                        self.log_manager.log(f"删除 {cls} 的奖惩记录: {ptype}分 {score}，备注: {note}")
                        break
        
        add_button = ttk.Button(button_frame, text="添加奖惩", command=add_punishment_local)
        add_button.pack(side=tk.LEFT, padx=5)
        delete_button = ttk.Button(button_frame, text="删除奖惩", command=delete_punishment_local)
        delete_button.pack(side=tk.LEFT, padx=5)
        
        list_frame = ttk.LabelFrame(main_frame, text="当前奖惩列表")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建Treeview和滚动条
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ("班级", "类型", "分值", "备注")
        local_vars['punishment_list_tree'] = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for col in columns:
            local_vars['punishment_list_tree'].heading(col, text=col)
            local_vars['punishment_list_tree'].column(col, width=150)
        
        # 正确配置滚动条
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=local_vars['punishment_list_tree'].yview)
        local_vars['punishment_list_tree'].configure(yscrollcommand=scrollbar.set)
        
        # 使用grid布局确保滚动条正确显示
        local_vars['punishment_list_tree'].grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        # 配置grid权重
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        for cls, punishments in self.punishments.items():
            for punishment in punishments:
                type_text = "奖励" if punishment["type"] == "add" else "惩罚"
                local_vars['punishment_list_tree'].insert("", "end", values=(cls, type_text, punishment["score"], punishment["note"]))
        
        # 窗口关闭时不需要清理引用，因为使用的是局部变量
        def on_closing():
            punishment_window.destroy()
        
        punishment_window.protocol("WM_DELETE_WINDOW", on_closing)


    

    
    def open_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("系统设置")
        settings_window.geometry("400x300")
        settings_window.transient(self.root)
        
        main_frame = ttk.Frame(settings_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        ttk.Button(main_frame, text="根目录设置", command=self.open_directory_settings, width=20).pack(pady=10)
        ttk.Button(main_frame, text="文件路径信息", command=self.open_file_info, width=20).pack(pady=10)
        ttk.Button(main_frame, text="项目分数设置", command=self.open_score_settings, width=20).pack(pady=10)
        ttk.Button(main_frame, text="班级管理", command=self.open_class_settings, width=20).pack(pady=10)
        
        ttk.Button(main_frame, text="关于", command=self.show_about, width=20).pack(pady=10)

    def open_directory_settings(self):
        dir_window = tk.Toplevel(self.root)
        dir_window.title("根目录设置")
        dir_window.geometry("500x200")
        dir_window.transient(self.root)
        
        main_frame = ttk.Frame(dir_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        ttk.Label(main_frame, text="当前根目录:").pack(anchor=tk.W)
        current_dir_label = ttk.Label(main_frame, text=self.settings_manager.get_root_directory(), 
                                     wraplength=400, justify=tk.LEFT)
        current_dir_label.pack(fill=tk.X, pady=(5, 10))
        
        def choose_directory():
            directory = filedialog.askdirectory(initialdir=self.settings_manager.get_root_directory())
            if directory:
                self.settings_manager.set_root_directory(directory)
                current_dir_label.config(text=directory)
                self.log_manager.log(f"更改根目录为: {directory}")
                messagebox.showinfo("设置", "根目录已更新")
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="选择根目录", command=choose_directory).pack(side=tk.LEFT)

    def open_file_info(self):
        info_window = tk.Toplevel(self.root)
        info_window.title("文件路径信息")
        info_window.geometry("500x150")
        info_window.transient(self.root)
        
        main_frame = ttk.Frame(info_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        ttk.Label(main_frame, text="数据文件路径:").pack(anchor=tk.W)
        data_file_label = ttk.Label(main_frame, text=self.settings_manager.get_data_file_path(), 
                                   wraplength=400, justify=tk.LEFT, foreground="blue")
        data_file_label.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(main_frame, text="日志文件路径:").pack(anchor=tk.W)
        log_file_label = ttk.Label(main_frame, text=self.settings_manager.get_log_file_path(), 
                                  wraplength=400, justify=tk.LEFT, foreground="blue")
        log_file_label.pack(fill=tk.X)

    def open_score_settings(self):
        scores_window = tk.Toplevel(self.root)
        scores_window.title("项目分数设置")
        scores_window.geometry("600x500")
        scores_window.transient(self.root)
        
        main_frame = ttk.Frame(scores_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        self.score_entries = {}
        for item_name in self.items:
            item_frame = ttk.Frame(scrollable_frame)
            item_frame.pack(fill=tk.X, padx=5, pady=2)
            
            ttk.Label(item_frame, text=item_name, width=15).pack(side=tk.LEFT)
            
            max_score = self.settings_manager.get_max_score(item_name)
            score_var = tk.StringVar(value=str(max_score))
            self.score_entries[item_name] = score_var
            
            score_entry = ttk.Entry(item_frame, textvariable=score_var, width=10)
            score_entry.pack(side=tk.LEFT, padx=5)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        save_frame = ttk.Frame(scores_window)
        save_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        def save_scores():
            try:
                for item_name, var in self.score_entries.items():
                    score = float(var.get())
                    self.settings_manager.set_max_score(item_name, score)
                
                self.reset_data()
                self.load_data()
                self.update_status("项目分数设置已保存并应用")
                self.log_manager.log("项目分数设置已更新")
                messagebox.showinfo("设置", "项目分数设置已保存并应用")
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字！")
        
        ttk.Button(save_frame, text="保存分数设置", command=save_scores).pack(side=tk.LEFT)

    def open_class_settings(self):
        class_window = tk.Toplevel(self.root)
        class_window.title("班级管理")
        class_window.geometry("800x400")
        class_window.transient(self.root)
        
        main_frame = ttk.Frame(class_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.create_class_management_section(main_frame)

    def create_class_management_section(self, parent):
        # 创建表格来显示班级信息
        columns = ("班级名称", "加权分数")
        self.class_tree = ttk.Treeview(parent, columns=columns, show="headings", height=10)
        
        # 设置表头
        for col in columns:
            self.class_tree.heading(col, text=col)
            self.class_tree.column(col, width=150)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.class_tree.yview)
        self.class_tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.class_tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 10))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 10))
        
        # 添加按钮框架
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=5)
        
        # 添加班级按钮
        ttk.Button(button_frame, text="添加班级", command=self.add_class).pack(side=tk.LEFT, padx=5)
        
        # 删除班级按钮
        ttk.Button(button_frame, text="删除选中班级", command=self.remove_class).pack(side=tk.LEFT, padx=5)
        
        # 上移班级按钮
        ttk.Button(button_frame, text="上移", command=self.move_class_up).pack(side=tk.LEFT, padx=5)
        
        # 下移班级按钮
        ttk.Button(button_frame, text="下移", command=self.move_class_down).pack(side=tk.LEFT, padx=5)
        
        # 保存设置按钮
        ttk.Button(button_frame, text="保存班级设置", command=self.save_class_settings).pack(side=tk.LEFT, padx=5)
        
        # 加载班级配置按钮
        ttk.Button(button_frame, text="加载班级配置", command=self.load_class_config_file).pack(side=tk.LEFT, padx=5)
        
        # 加载现有班级数据
        self.load_class_data_to_tree()
        
        # 绑定选择事件
        self.class_tree.bind("<<TreeviewSelect>>", self.on_class_select)
        # 绑定双击事件用于编辑
        self.class_tree.bind("<Double-1>", self.on_class_tree_double_click)
        
        # 保存选中的项目
        self.selected_class_item = None
    
    def on_class_tree_double_click(self, event):
        """处理表格双击事件，允许编辑单元格"""
        # 获取点击的项目和列
        item = self.class_tree.identify_row(event.y)
        column = self.class_tree.identify_column(event.x)
        
        if item and column:
            # 获取当前值
            values = self.class_tree.item(item, "values")
            col_index = int(column[1:]) - 1  # 转换为0基索引
            current_value = values[col_index]
            
            # 创建编辑窗口
            self.create_edit_window(item, col_index, current_value)
    
    def create_edit_window(self, item, col_index, current_value):
        """创建编辑窗口"""
        # 创建顶层窗口
        edit_window = tk.Toplevel(self.class_tree)
        edit_window.title("编辑")
        edit_window.geometry("300x100")
        edit_window.transient(self.class_tree)
        edit_window.grab_set()
        
        # 居中显示
        edit_window.geometry("+{}+{}".format(
            edit_window.winfo_screenwidth() // 2 - 150,
            edit_window.winfo_screenheight() // 2 - 50
        ))
        
        # 创建输入框
        ttk.Label(edit_window, text="请输入新值:").pack(pady=5)
        entry_var = tk.StringVar(value=current_value)
        entry = ttk.Entry(edit_window, textvariable=entry_var, width=30)
        entry.pack(pady=5)
        entry.select_range(0, tk.END)
        entry.focus()
        
        # 保存按钮
        def save_edit():
            new_value = entry_var.get()
            values = list(self.class_tree.item(item, "values"))
            values[col_index] = new_value
            self.class_tree.item(item, values=values)
            edit_window.destroy()
        
        # 按钮框架
        button_frame = ttk.Frame(edit_window)
        button_frame.pack(pady=5)
        
        ttk.Button(button_frame, text="保存", command=save_edit).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=edit_window.destroy).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键保存
        entry.bind("<Return>", lambda e: save_edit())
        
        # 等待窗口关闭
        edit_window.wait_window()
    
    def load_class_data_to_tree(self):
        """加载现有班级数据到表格"""
        # 清空现有数据
        for item in self.class_tree.get_children():
            self.class_tree.delete(item)
        
        # 获取现有班级数据
        current_classes = self.settings_manager.get_classes()
        weighted_addition = self.settings_manager.get_weighted_addition()
        
        # 添加到表格
        for class_name in current_classes:
            weighted_score = weighted_addition.get(class_name, 0)
            self.class_tree.insert("", "end", values=(class_name, str(weighted_score)))
    
    def on_class_select(self, event):
        """处理班级选择事件"""
        selected_items = self.class_tree.selection()
        if selected_items:
            self.selected_class_item = selected_items[0]
        else:
            self.selected_class_item = None
    
    def add_class(self):
        # 弹出对话框让用户输入班级名称
        new_class_name = tk.simpledialog.askstring("添加班级", "请输入班级名称:")
        if not new_class_name:
            return  # 用户取消了输入
        
        new_weighted_score = "0"
        
        # 检查是否已存在同名班级
        existing_classes = [self.class_tree.item(item, "values")[0] for item in self.class_tree.get_children()]
        if new_class_name in existing_classes:
            messagebox.showwarning("警告", f"班级 '{new_class_name}' 已存在！")
            return
        
        # 添加到表格
        self.class_tree.insert("", "end", values=(new_class_name, new_weighted_score))
    
    def remove_class(self):
        # 删除选中的班级
        selected_items = self.class_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的班级！")
            return
        
        for item in selected_items:
            self.class_tree.delete(item)
        
        # 保存更改
        self.save_class_settings()
    
    def move_class_up(self):
        """上移选中的班级"""
        selected = self.class_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要移动的班级！")
            return
        
        # 获取所有项目和选中项目的索引
        all_items = self.class_tree.get_children()
        selected_item = selected[0]
        current_index = all_items.index(selected_item)
        
        # 如果已经在最顶部，则不移动
        if current_index == 0:
            return
        
        # 获取要交换位置的项目
        prev_item = all_items[current_index - 1]
        
        # 获取两个项目的数据
        selected_values = self.class_tree.item(selected_item, "values")
        prev_values = self.class_tree.item(prev_item, "values")
        
        # 交换显示位置
        self.class_tree.move(selected_item, "", current_index - 1)
        
        # 重新选择移动的项目
        self.class_tree.selection_set(selected_item)
    
    def move_class_down(self):
        """下移选中的班级"""
        selected = self.class_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要移动的班级！")
            return
        
        # 获取所有项目和选中项目的索引
        all_items = self.class_tree.get_children()
        selected_item = selected[0]
        current_index = all_items.index(selected_item)
        
        # 如果已经在最底部，则不移动
        if current_index == len(all_items) - 1:
            return
        
        # 获取要交换位置的项目
        next_item = all_items[current_index + 1]
        
        # 获取两个项目的数据
        selected_values = self.class_tree.item(selected_item, "values")
        next_values = self.class_tree.item(next_item, "values")
        
        # 交换显示位置
        self.class_tree.move(selected_item, "", current_index + 1)
        
        # 重新选择移动的项目
        self.class_tree.selection_set(selected_item)
    
    def save_class_settings(self):
        try:
            classes = []
            weighted_addition = {}
            
            # 从表格中获取班级数据，保持顺序
            for item in self.class_tree.get_children():
                values = self.class_tree.item(item, "values")
                class_name = values[0].strip()
                if class_name:
                    classes.append(class_name)
                    
                    # 获取加权分数
                    try:
                        weighted_value = float(values[1])
                    except ValueError:
                        weighted_value = 0
                    weighted_addition[class_name] = weighted_value
            
            self.settings_manager.set_classes(classes)
            self.settings_manager.set_weighted_addition(weighted_addition)
            
            self.classes = classes
            self.weighted_addition = weighted_addition.copy()
            
            # 更新主窗口的班级下拉列表
            if self.class_combobox is not None:
                self.class_combobox['values'] = self.classes
            
            # 保存到班级配置文件
            self.save_class_config_file()
            
            messagebox.showinfo("设置", "班级设置已保存")
            self.log_manager.log("班级设置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存班级设置时出错：{str(e)}")
            self.log_manager.log(f"保存班级设置时出错: {str(e)}")
    
    def save_class_config_file(self):
        """保存班级配置到独立的JSON文件"""
        try:
            class_config = {
                "classes": self.classes,
                "weighted_addition": self.weighted_addition
            }
            
            # 使用固定文件名保存班级配置
            config_file = os.path.join(self.settings_manager.get_root_directory(), "class_config.json")
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(class_config, f, ensure_ascii=False, indent=2, default=str)
            
            self.log_manager.log(f"班级配置已保存到 {config_file}")
        except Exception as e:
            messagebox.showerror("错误", f"保存班级配置文件时出错：{str(e)}")
            self.log_manager.log(f"保存班级配置文件时出错: {str(e)}")
    
    def load_class_config_file(self):
        """从独立的JSON文件加载班级配置"""
        try:
            config_file = os.path.join(self.settings_manager.get_root_directory(), "class_config.json")
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    class_config = json.load(f)
                
                # 更新班级和加权分数
                self.classes = class_config.get("classes", [])
                self.weighted_addition = class_config.get("weighted_addition", {})
                
                # 更新设置管理器
                self.settings_manager.set_classes(self.classes)
                self.settings_manager.set_weighted_addition(self.weighted_addition)
                
                # 更新表格显示
                self.update_class_tree_from_config()
                
                # 更新主窗口的班级下拉列表
                if self.class_combobox is not None:
                    self.class_combobox['values'] = self.classes
                
                self.log_manager.log(f"班级配置已从 {config_file} 加载")
                return True
            else:
                self.log_manager.log("班级配置文件不存在，使用默认配置")
                return False
        except Exception as e:
            messagebox.showerror("错误", f"加载班级配置文件时出错：{str(e)}")
            self.log_manager.log(f"加载班级配置文件时出错: {str(e)}")
            return False
    
    def update_class_tree_from_config(self):
        """根据配置更新班级表格显示"""
        # 清空现有数据
        for item in self.class_tree.get_children():
            self.class_tree.delete(item)
        
        # 添加班级数据到表格
        for class_name in self.classes:
            weighted_score = self.weighted_addition.get(class_name, 0)
            self.class_tree.insert("", "end", values=(class_name, weighted_score))
    
    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("关于流动红旗评比系统")
        about_window.geometry("720x650")
        about_window.resizable(True, True)
        about_window.configure(bg="#f5f5f5")
        
        about_window.transient(self.root)
        about_window.grab_set()
        
        # 创建主框架
        main_frame = ttk.Frame(about_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题区域
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(
            title_frame, 
            text="流动红旗评比系统", 
            font=("微软雅黑", 20, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            title_frame,
            text="学校流动红旗评比管理系统",
            font=("微软雅黑", 10),
            foreground="#7f8c8d"
        )
        subtitle_label.pack(pady=(5, 0))
        
        # 内容区域
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧信息区域
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 版本信息
        version_frame = ttk.LabelFrame(left_frame, text="版本信息", padding="10")
        version_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(version_frame, text="版本: V1.5.0", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W)
        ttk.Label(version_frame, text="发布日期: 2025年", font=("微软雅黑", 10)).pack(anchor=tk.W, pady=(5, 0))
        
        # 开发者信息
        dev_frame = ttk.LabelFrame(left_frame, text="开发者信息", padding="10")
        dev_frame.pack(fill=tk.X, pady=(0, 15))
        
        github_link = "https://github.com/Bao-Jiaozixing/flowing-red-flag-evaluation"
        link_label = ttk.Label(dev_frame, text=f"开发团队: {github_link}", font=("微软雅黑", 10), foreground="blue", cursor="hand2")
        link_label.pack(anchor=tk.W)
        link_label.bind("<Button-1>", lambda e: self.open_link(github_link))
        
        # 系统功能
        info_frame = ttk.LabelFrame(left_frame, text="系统功能", padding="10")
        info_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(info_frame, text="支持功能:", font=("微软雅黑", 10, "bold")).pack(anchor=tk.W)
        
        features = [
            "• 日常评分管理", "• 惩罚加分管理", "• 数据导入导出",
            "• 自动计算总分", "• 班级排名显示", "• 历史记录保存"
        ]
        
        for feature in features:
            ttk.Label(info_frame, text=feature, font=("微软雅黑", 9)).pack(anchor=tk.W, padx=(10, 0), pady=2)
        
        # 右侧信息区域
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # 使用模型
        model_frame = ttk.LabelFrame(right_frame, text="使用模型", padding="10")
        model_frame.pack(fill=tk.X, pady=(0, 15))
        
        models = [
            "• Deepseek-V3",
            "• Deepseek-R1",
            "• Qwen-3-Coder"
        ]
        
        for model in models:
            ttk.Label(model_frame, text=model, font=("微软雅黑", 9)).pack(anchor=tk.W, pady=1)
        
        # 第三方库声明
        license_frame = ttk.LabelFrame(right_frame, text="第三方库声明", padding="10")
        license_frame.pack(fill=tk.BOTH, expand=True)
        
        libraries = [
            "• tkinter - Python标准GUI库",
            "• pandas - 数据处理库 (用于Excel导出)",
            "• openpyxl - Excel文件处理库 (pandas依赖)"
        ]
        
        for library in libraries:
            ttk.Label(license_frame, text=library, font=("微软雅黑", 9)).pack(anchor=tk.W, padx=(0, 0), pady=1)
        
        # 底部版权信息
        copyright_frame = ttk.Frame(main_frame)
        copyright_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Separator(copyright_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            copyright_frame, 
            text="© 2025 流动红旗评比系统. 保留所有权利.", 
            font=("微软雅黑", 9), 
            foreground="#95a5a6"
        ).pack()
        
        # 关闭按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        close_button = ttk.Button(
            button_frame, 
            text="关闭", 
            command=about_window.destroy,
            width=15
        )
        close_button.pack(side=tk.RIGHT)

    def open_link(self, url):
        import webbrowser
        webbrowser.open_new(url)

    def save_data(self):
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"流动红旗分数数据_{timestamp}.json"
            
            file_path = filedialog.asksaveasfilename(
                initialfile=default_filename,
                defaultextension=".json",
                filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
                title="保存数据",
                initialdir=self.settings_manager.get_root_directory()
            )
            
            if not file_path:
                return
            
            self.log_manager.log(f"用户选择的文件路径: {file_path}")
            
            if not file_path.strip():
                error_msg = "文件路径不能为空"
                self.root.lift()
                self.root.focus_force()
                messagebox.showerror("错误", error_msg)
                self.log_manager.log(f"保存数据时出错: {error_msg}")
                return
            
            file_path = os.path.normpath(file_path)
            self.log_manager.log(f"规范化后的文件路径: {file_path}")
            
            directory = os.path.dirname(file_path)
            self.log_manager.log(f"目录路径: {directory}")
            
            if directory:
                if not os.path.exists(directory):
                    self.log_manager.log(f"目录不存在，尝试创建目录: {directory}")
                    os.makedirs(directory, exist_ok=True)
                    self.log_manager.log(f"目录创建成功: {directory}")
                elif not os.access(directory, os.W_OK):
                    error_msg = f"目录没有写入权限: {directory}"
                    self.root.lift()
                    self.root.focus_force()
                    messagebox.showerror("权限错误", error_msg)
                    self.log_manager.log(f"保存数据时出错: {error_msg}")
                    return
            
            if os.path.exists(file_path) and not os.access(file_path, os.W_OK):
                error_msg = f"文件没有写入权限: {file_path}"
                self.root.lift()
                self.root.focus_force()
                messagebox.showerror("权限错误", error_msg)
                self.log_manager.log(f"保存数据时出错: {error_msg}")
                return
            
            data = {
                'scores': self.get_all_scores(),
                'punishments': self.punishments,
                'weighted_addition': self.weighted_addition,
                'classes': self.settings_manager.get_classes(),
                'save_time': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            self.save_snapshot()
            
            self.update_status("数据保存成功")
            self.log_manager.log("数据保存成功")
            messagebox.showinfo("成功", f"数据已保存到 {file_path}")
        except Exception as e:
            error_msg = f"保存数据时出错：{str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("错误", error_msg)
            self.log_manager.log(f"保存数据时出错: {error_msg}")
    
    def save_as_data(self):
        try:
            data = {
                'scores': self.get_all_scores(),
                'punishments': self.punishments,
                'weighted_addition': self.weighted_addition,
                'classes': self.settings_manager.get_classes(),
                'save_time': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
                title="另存为"
            )
            
            if file_path:
                self.log_manager.log(f"用户选择的文件路径: {file_path}")
                
                if not file_path.strip():
                    error_msg = "文件路径不能为空"
                    self.root.lift()
                    self.root.focus_force()
                    messagebox.showerror("错误", error_msg)
                    self.log_manager.log(f"另存为数据时出错: {error_msg}")
                    return
                
                file_path = os.path.normpath(file_path)
                self.log_manager.log(f"规范化后的文件路径: {file_path}")
                
                directory = os.path.dirname(file_path)
                self.log_manager.log(f"目录路径: {directory}")
                
                if directory:
                    if not os.path.exists(directory):
                        self.log_manager.log(f"目录不存在，尝试创建目录: {directory}")
                        os.makedirs(directory, exist_ok=True)
                        self.log_manager.log(f"目录创建成功: {directory}")
                    elif not os.access(directory, os.W_OK):
                        error_msg = f"目录没有写入权限: {directory}"
                        self.root.lift()
                        self.root.focus_force()
                        messagebox.showerror("权限错误", error_msg)
                        self.log_manager.log(f"另存为数据时出错: {error_msg}")
                        return
                
                if os.path.exists(file_path) and not os.access(file_path, os.W_OK):
                    error_msg = f"文件没有写入权限: {file_path}"
                    self.root.lift()
                    self.root.focus_force()
                    messagebox.showerror("权限错误", error_msg)
                    self.log_manager.log(f"另存为数据时出错: {error_msg}")
                    return
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                
                self.history_manager.add_record(data)
                self.update_status(f"数据已另存为 {file_path}")
                self.log_manager.log(f"数据另存为: {file_path}")
                messagebox.showinfo("成功", f"数据已另存为 {file_path}")
        except PermissionError as e:
            error_msg = f"没有权限保存到指定位置，请选择其他位置或以管理员身份运行程序: {str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("权限错误", error_msg)
            self.log_manager.log(f"另存为数据时出错: {error_msg}")
        except FileNotFoundError as e:
            error_msg = f"指定的文件路径不存在，请检查路径是否正确: {str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("路径错误", error_msg)
            self.log_manager.log(f"另存为数据时出错: {error_msg}")
        except OSError as e:
            error_msg = f"操作系统错误: {str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("系统错误", error_msg)
            self.log_manager.log(f"另存为数据时出错: {error_msg}")
        except Exception as e:
            error_msg = f"另存为数据时出错：{str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("错误", error_msg)
            self.log_manager.log(f"另存为数据时出错: {error_msg}")
    
    def save_total_score_table(self):
        try:
            self.calculate_totals()
            
            data = []
            columns = ["班级", "早迟到", "早读", "节能开窗", "仪容仪表", "间操", "午休", 
                      "上下午各班卫生", "巡视", "及时上交文件", "宿舍", "加权", "惩罚", "惩罚备注", "总分"]
            
            for cls in self.classes:
                row = [cls]
                for page_name in self.items:
                    max_score = self.settings_manager.get_max_score(page_name)
                    
                    if page_name in self.dual_period_items:
                        am_avg = self.get_class_average(self.pages[page_name].am_tree, cls)
                        pm_avg = self.get_class_average(self.pages[page_name].pm_tree, cls)
                        avg_score = (am_avg + pm_avg) / 2
                        row.append(round(avg_score, 2))
                    else:
                        for widget in self.pages[page_name].winfo_children():
                            if isinstance(widget, ttk.Treeview):
                                for item in widget.get_children():
                                    item_values = widget.item(item, "values")
                                    if item_values[0] == cls:
                                        row.append(item_values[6])
                                        break
                row.append(self.weighted_addition[cls])
                punishment_score = 0
                punishment_notes = []
                if cls in self.punishments:
                    for punishment in self.punishments[cls]:
                        if punishment["type"] == "add":
                            punishment_score += punishment["score"]
                        elif punishment["type"] == "subtract":
                            punishment_score -= punishment["score"]
                        punishment_notes.append(punishment["note"])
                row.append(punishment_score)
                row.append("\n".join(punishment_notes))
                total_score = sum(map(float, row[1:-2])) + row[-2]
                row.append(round(total_score, 2))
                
                data.append(row)
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv"), ("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
                title="保存总分表格"
            )
            
            if file_path:
                self.log_manager.log(f"用户选择的文件路径: {file_path}")
                
                if not file_path.strip():
                    error_msg = "文件路径不能为空"
                    self.root.lift()
                    self.root.focus_force()
                    messagebox.showerror("错误", error_msg)
                    self.log_manager.log(f"保存总分表格时出错: {error_msg}")
                    return
                
                file_path = os.path.normpath(file_path)
                self.log_manager.log(f"规范化后的文件路径: {file_path}")
                
                directory = os.path.dirname(file_path)
                self.log_manager.log(f"目录路径: {directory}")
                
                if directory:
                    if not os.path.exists(directory):
                        self.log_manager.log(f"目录不存在，尝试创建目录: {directory}")
                        os.makedirs(directory, exist_ok=True)
                        self.log_manager.log(f"目录创建成功: {directory}")
                    elif not os.access(directory, os.W_OK):
                        error_msg = f"目录没有写入权限: {directory}"
                        self.root.lift()
                        self.root.focus_force()
                        messagebox.showerror("权限错误", error_msg)
                        self.log_manager.log(f"保存总分表格时出错: {error_msg}")
                        return
                
            if os.path.exists(file_path) and not os.access(file_path, os.W_OK):
                error_msg = f"文件没有写入权限: {file_path}"
                self.root.lift()
                self.root.focus_force()
                messagebox.showerror("权限错误", error_msg)
                self.log_manager.log(f"保存总分表格时出错: {error_msg}")
                return
                
                if file_path.endswith('.csv'):
                    import csv
                    with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                        writer = csv.writer(f)
                        writer.writerow(columns)
                        writer.writerows(data)
                elif file_path.endswith('.xlsx'):
                    try:
                        import pandas as pd
                        df = pd.DataFrame(data, columns=columns)
                        df.to_excel(file_path, index=False)
                    except ImportError:
                        error_msg = "未安装pandas库，无法导出Excel\n请运行: pip install pandas openpyxl"
                        self.root.lift()
                        self.root.focus_force()
                        messagebox.showerror("错误", error_msg)
                        self.log_manager.log(f"保存总分表格时出错: {error_msg}")
                        return
                else:
                    import csv
                    with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                        writer = csv.writer(f)
                        writer.writerow(columns)
                        writer.writerows(data)
                
                self.update_status(f"总分表格已保存到 {file_path}")
                self.log_manager.log(f"总分表格已保存到: {file_path}")
                messagebox.showinfo("成功", f"总分表格已保存到 {file_path}")
        except PermissionError as e:
            error_msg = f"没有权限保存到指定位置，请选择其他位置或以管理员身份运行程序: {str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("权限错误", error_msg)
            self.log_manager.log(f"保存总分表格时出错: {error_msg}")
        except FileNotFoundError as e:
            error_msg = f"指定的文件路径不存在，请检查路径是否正确: {str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("路径错误", error_msg)
            self.log_manager.log(f"保存总分表格时出错: {error_msg}")
        except OSError as e:
            error_msg = f"操作系统错误: {str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("系统错误", error_msg)
            self.log_manager.log(f"保存总分表格时出错: {error_msg}")
        except Exception as e:
            error_msg = f"保存总分表格时出错：{str(e)}"
            self.root.lift()
            self.root.focus_force()
            messagebox.showerror("错误", error_msg)
            self.log_manager.log(f"保存总分表格时出错: {error_msg}")
    
    def on_closing(self):
        if hasattr(self, 'time_after_id'):
            self.root.after_cancel(self.time_after_id)
        self.root.destroy()
    
    def load_data(self):
        try:
            data_file = filedialog.askopenfilename(
                defaultextension=".json",
                filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
                title="选择要加载的数据文件"
            )
            
            if not data_file:
                return
            
            with open(data_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            # 直接同步班级设置
            loaded_classes = data.get('classes', self.settings_manager.get_classes())
            loaded_scores = data.get('scores', {})
            
            # 同步班级设置
            self.settings_manager.set_classes(loaded_classes)
            self.classes = loaded_classes
            
            # 更新加权分数设置以匹配新班级
            loaded_weighted_addition = data.get('weighted_addition', {})
            for cls in loaded_classes:
                if cls not in loaded_weighted_addition:
                    loaded_weighted_addition[cls] = 0  # 默认加权分数为0
            self.settings_manager.set_weighted_addition(loaded_weighted_addition)
            self.weighted_addition = loaded_weighted_addition.copy()
                
            self.load_scores(loaded_scores)
            self.punishments = data.get('punishments', {})
            weighted_data = data.get('weighted_addition', {})
            self.weighted_addition.update(weighted_data)
            
            # self.classes已在此前设置为loaded_classes，无需重复设置
            # self.classes = loaded_classes
            
            self.update_status("数据加载成功")
            self.log_manager.log("数据加载成功")
            messagebox.showinfo("成功", f"数据已从 {data_file} 加载")
        except FileNotFoundError:
            self.update_status("未找到保存的数据文件，使用默认数据")
            self.log_manager.log("未找到保存的数据文件，使用默认数据")
        except Exception as e:
            error_msg = f"加载数据时出错：{str(e)}"
            messagebox.showerror("错误", error_msg)
            self.log_manager.log(f"加载数据时出错: {error_msg}")
    
    def get_all_scores(self):
        scores = {}
        for page_name, page_frame in self.pages.items():
            if page_name in self.dual_period_items:
                scores[f"{page_name}_am"] = self.get_tree_data(page_frame.am_tree)
                scores[f"{page_name}_pm"] = self.get_tree_data(page_frame.pm_tree)
            else:
                for widget in page_frame.winfo_children():
                    if isinstance(widget, ttk.Treeview):
                        scores[page_name] = self.get_tree_data(widget)
                        break
        return scores
    
    def get_tree_data(self, tree):
        return [tree.item(item, "values") for item in tree.get_children()]
    
    def save_snapshot(self):
        snapshot = {
            'scores': self.get_all_scores(),
            'punishments': self.punishments.copy(),
            'weighted_addition': self.weighted_addition.copy()
        }
        self.history_manager.add_record(snapshot)
    
    def load_scores(self, scores_data):
        for page_name, page_data in scores_data.items():
            if page_name.endswith("_am") and len(page_name) > 3:
                real_page_name = page_name[:-3]
                if real_page_name in self.dual_period_items and real_page_name in self.pages:
                    self.set_tree_data(self.pages[real_page_name].am_tree, page_data)
            elif page_name.endswith("_pm") and len(page_name) > 3:
                real_page_name = page_name[:-3]
                if real_page_name in self.dual_period_items and real_page_name in self.pages:
                    self.set_tree_data(self.pages[real_page_name].pm_tree, page_data)
            elif page_name in self.pages:
                for widget in self.pages[page_name].winfo_children():
                    if isinstance(widget, ttk.Treeview):
                        self.set_tree_data(widget, page_data)
                        break
    
    def set_tree_data(self, tree, data):
        for i, item_data in enumerate(data):
            item = tree.get_children()[i] if i < len(tree.get_children()) else None
            if item:
                tree.item(item, values=item_data)
    


    
    def show_history(self):
        history_window = tk.Toplevel(self.root)
        history_window.title("历史记录")
        history_window.geometry("800x500")
        history_window.transient(self.root)
        
        main_frame = ttk.Frame(history_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ("保存时间", "操作")
        tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        tree.heading("保存时间", text="保存时间")
        tree.heading("操作", text="操作")
        tree.column("保存时间", width=200)
        tree.column("操作", width=300)
        
        tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)
        
        history = self.history_manager.get_history()
        for record in history:
            tree.insert("", "end", values=(record['date'], "保存数据"))
        
        load_frame = ttk.Frame(history_window)
        load_frame.pack(fill=tk.X, padx=10, pady=5)
        
        def load_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showerror("错误", "请选择一条记录")
                return
            
            item = selected[0]
            index = tree.index(item)
            
            history = self.history_manager.get_history()
            if 0 <= index < len(history):
                try:
                    record_data = history[index]['data']
                    self.load_scores(record_data.get('scores', {}))
                    self.punishments = record_data.get('punishments', {}).copy()
                    weighted_data = record_data.get('weighted_addition', {})
                    self.weighted_addition.clear()
                    self.weighted_addition.update(weighted_data)
                    self.update_status(f"已加载历史记录: {history[index]['date']}")
                    self.log_manager.log(f"加载历史记录: {history[index]['date']}")
                    self.calculate_totals()  # 重新计算总分
                    history_window.destroy()
                except Exception as e:
                    error_msg = f"加载历史记录时出错：{str(e)}"
                    messagebox.showerror("错误", error_msg)
                    self.log_manager.log(error_msg)
            else:
                messagebox.showerror("错误", f"无法加载选中的记录：索引 {index} 超出范围 [0, {len(history)})")
        
        ttk.Button(load_frame, text="加载选中记录", command=load_selected).pack(side=tk.LEFT)
    
    def undo_action(self):
        if self.history_manager.can_undo():
            previous_data = self.history_manager.undo()
            if previous_data:
                self.load_scores(previous_data.get('scores', {}))
                self.punishments = previous_data.get('punishments', {}).copy()
                weighted_data = previous_data.get('weighted_addition', {})
                self.weighted_addition.clear()
                self.weighted_addition.update(weighted_data)
                self.update_status("已撤销操作")
                self.log_manager.log("执行撤销操作")
                self.calculate_totals()
            else:
                self.update_status("无法撤销操作")
        else:
            self.update_status("没有可撤销的操作")
            self.log_manager.log("尝试撤销操作但没有历史记录")
    
    def redo_action(self):
        if self.history_manager.can_redo():
            next_data = self.history_manager.redo()
            if next_data:
                self.load_scores(next_data.get('scores', {}))
                self.punishments = next_data.get('punishments', {}).copy()
                weighted_data = next_data.get('weighted_addition', {})
                self.weighted_addition.clear()
                self.weighted_addition.update(weighted_data)
                self.update_status("已重做操作")
                self.log_manager.log("执行重做操作")
                self.calculate_totals()
            else:
                self.update_status("无法重做操作")
        else:
            self.update_status("没有可重做的操作")
            self.log_manager.log("尝试重做操作但没有历史记录")
    
    def toggle_fullscreen(self):
        self.root.attributes('-fullscreen', not self.root.attributes('-fullscreen'))
    
    def on_tab_change(self, event):
        self.destroy_entry()
    
    def on_click_anywhere(self, event):
        self.destroy_entry()
    
    def destroy_entry(self):
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Entry) and widget.winfo_viewable():
                widget.destroy()

if __name__ == "__main__":
    root = tk.Tk()  # 使用标准的tk.Tk
    
    try:
        root.iconbitmap('icon.ico')
    except tk.TclError:
        print("无法加载图标文件，请确保icon.ico文件存在于程序目录中")
    
    if chinese_font:
        root.option_add("*Font", (chinese_font, 9))
    
    app = FlowingRedFlagEvaluationSystem(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()