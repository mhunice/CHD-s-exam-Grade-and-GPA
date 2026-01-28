import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import pandas as pd
import time
from io import StringIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.driver_cache import DriverCacheManager

class GradeCrawlerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("教务系统绩点计算器")
        self.root.geometry("900x700")

        # --- 配置区域 ---
        self.LOGIN_URL = "http://bkjw.chd.edu.cn/eams/home.action"
        self.GRADE_URL = "http://bkjw.chd.edu.cn/eams/teach/grade/course/person.action"
        self.default_excludes = "科学探索与技术创新（2022）|经典阅读与写作沟通|社会科学与公共责任"

        # UI 初始化
        self.create_widgets()

        # 预填默认账号
        self.entry_user.insert(0, "")
        self.entry_pwd.insert(0, "")

    def create_widgets(self):
        # 1. 顶部控制面板
        frame_top = ttk.LabelFrame(self.root, text="登录设置", padding=10)
        frame_top.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_top, text="学号:").grid(row=0, column=0, padx=5, sticky="e")
        self.entry_user = ttk.Entry(frame_top, width=20)
        self.entry_user.grid(row=0, column=1, padx=5)

        ttk.Label(frame_top, text="密码:").grid(row=0, column=2, padx=5, sticky="e")
        self.entry_pwd = ttk.Entry(frame_top, width=20, show="*")
        self.entry_pwd.grid(row=0, column=3, padx=5)

        self.btn_start = ttk.Button(frame_top, text="开始爬取并计算", command=self.start_thread)
        self.btn_start.grid(row=0, column=4, padx=20)

        # 2. 排除设置
        frame_exclude = ttk.LabelFrame(self.root, text="排除课程关键词 (用 | 分隔)", padding=10)
        frame_exclude.pack(fill="x", padx=10, pady=5)

        self.entry_exclude = ttk.Entry(frame_exclude)
        self.entry_exclude.pack(fill="x", padx=5)
        self.entry_exclude.insert(0, self.default_excludes)

        # 3. 日志区域
        frame_log = ttk.LabelFrame(self.root, text="运行日志", padding=10)
        frame_log.pack(fill="x", padx=10, pady=5)

        self.text_log = scrolledtext.ScrolledText(frame_log, height=8, state='disabled', font=("Consolas", 9))
        self.text_log.pack(fill="both", expand=True)

        # 4. 结果展示区域 (Treeview)
        frame_result = ttk.LabelFrame(self.root, text="成绩明细与结果", padding=10)
        frame_result.pack(fill="both", expand=True, padx=10, pady=5)

        self.lbl_result = ttk.Label(frame_result, text="等待计算...", font=("Arial", 12, "bold"), foreground="blue")
        self.lbl_result.pack(pady=5)

        columns = ("课程名称", "课程类别", "学分", "绩点", "状态")
        self.tree = ttk.Treeview(frame_result, columns=columns, show="headings", height=15)

        self.tree.column("课程名称", width=250)
        self.tree.column("课程类别", width=150)
        self.tree.column("学分", width=60, anchor="center")
        self.tree.column("绩点", width=60, anchor="center")
        self.tree.column("状态", width=80, anchor="center")

        for col in columns:
            self.tree.heading(col, text=col)

        scrollbar = ttk.Scrollbar(frame_result, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def log(self, message):
        """线程安全的日志打印"""
        self.root.after(0, self._append_log, message)

    def _append_log(self, message):
        self.text_log.config(state='normal')
        self.text_log.insert(tk.END, message + "\n")
        self.text_log.see(tk.END)
        self.text_log.config(state='disabled')

    def _safe_insert(self, values):
        """线程安全的表格插入辅助函数"""
        self.tree.insert("", "end", values=values)

    def start_thread(self):
        """启动后台线程"""
        username = self.entry_user.get()
        password = self.entry_pwd.get()
        excludes = self.entry_exclude.get().split("|")

        if not username or not password:
            messagebox.showwarning("提示", "请输入学号和密码")
            return

        self.btn_start.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.lbl_result.config(text="正在运行中...")

        thread = threading.Thread(target=self.run_crawler, args=(username, password, excludes))
        thread.daemon = True
        thread.start()

    # ==========================
    # 核心逻辑
    # ==========================

    def _login_logic(self, driver, username, password):
        wait = WebDriverWait(driver, 20)
        self.log(f"正在打开登录页: {self.LOGIN_URL}")
        driver.get(self.LOGIN_URL)

        try:
            user = wait.until(EC.presence_of_element_located((By.ID, "username")))
            pwd = wait.until(EC.presence_of_element_located((By.ID, "password")))

            user.clear()
            user.send_keys(username)
            self.log("✔ 学号已输入")

            # DOM 注入密码
            driver.execute_script("arguments[0].value = arguments[1];", pwd, password)
            driver.execute_script("""
                arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
                arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
            """, pwd)

            self.log("✔ 密码已写入（DOM 模式）")
            pwd.send_keys(Keys.ENTER)

            time.sleep(3)

            if "login" in driver.current_url.lower():
                self.log("⚠ 自动登录未跳转，可能有验证码")
                messagebox.showinfo("提示", "如有验证码，请在浏览器中手动完成登录\n完成后点击【确定】继续")
                self.log("用户已确认手动登录完成，继续执行...")
            else:
                self.log("🎉 自动登录成功")

        except Exception as e:
            self.log(f"登录过程出错: {e}")
            raise e



    def run_crawler(self, username, password, exclude_list):
        driver = None
        try:
            self.log("🚀 正在初始化浏览器（自动匹配驱动+静默模式）...")

            # 1. 设置静默模式参数
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')  # 开启静默模式
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080')
            options.add_argument('--no-sandbox')  # 增加稳定性
            options.add_argument('--disable-dev-shm-usage')
            # 伪装 UA，防止因 headless 被反爬拦截
            options.add_argument(
                '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

            # 2. 自动下载/匹配 ChromeDriver 并启动
            os.environ['WDM_CDN_URL'] = ""

            # 使用 Service 包装
            from selenium.webdriver.chrome.service import Service

            # 自动下载并指定镜像源
            driver_path = ChromeDriverManager(url="").install()
            service_obj = Service(driver_path)

            driver = webdriver.Chrome(service=service_obj, options=options)
            self._login_logic(driver, username, password)

            self.log("正在跳转至成绩页...")
            driver.get(self.GRADE_URL)

            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "gridtable")))
            self.log("✅ 获取到成绩表格数据")

            html = driver.page_source
            self.process_data(html, exclude_list)

        except Exception as e:
            self.log(f"❌ 发生错误: {str(e)}")
            messagebox.showerror("错误", f"运行出错:\n{str(e)}")
        finally:
            self.root.after(0, lambda: self.btn_start.config(state="normal"))
            if driver:
                self.log("关闭浏览器")
                driver.quit()

    def process_data(self, html, exclude_list):
        """解析 HTML 并计算绩点"""
        try:
            tables = pd.read_html(StringIO(html))
            grade_df = None

            for table in tables:
                cols = table.columns.astype(str)
                if any("课程" in c for c in cols) and any("绩点" in c for c in cols):
                    grade_df = table
                    break

            if grade_df is None:
                raise Exception("未找到包含'课程'和'绩点'列的表格")

            df = grade_df.copy()
            df.columns = df.columns.astype(str)

            col_map = {}
            for c in df.columns:
                if "课程" in c and "类别" not in c:
                    col_map["name"] = c
                elif "类别" in c:
                    col_map["type"] = c
                elif "学分" in c:
                    col_map["credit"] = c
                elif "绩点" in c:
                    col_map["gpa"] = c

            clean_df = df[[col_map["name"], col_map["type"], col_map["credit"], col_map["gpa"]]].copy()
            clean_df.columns = ["课程名称", "课程类别", "学分", "绩点"]

            clean_df["学分"] = pd.to_numeric(clean_df["学分"], errors="coerce")
            clean_df["绩点"] = pd.to_numeric(clean_df["绩点"], errors="coerce")
            clean_df = clean_df.dropna(subset=["学分", "绩点"])

            total_credit = 0
            total_score = 0

            import re
            pattern = "|".join(exclude_list) if exclude_list else "____ImpossibleMatch____"

            for index, row in clean_df.iterrows():
                name = str(row["课程名称"])
                ctype = str(row["课程类别"])
                credit = float(row["学分"])
                gpa = float(row["绩点"])

                is_excluded = False
                if exclude_list:
                    if re.search(pattern, ctype) or re.search(pattern, name):
                        is_excluded = True

                status = "排除" if is_excluded else "计入"

                # --- 修复点：使用 _safe_insert 避免关键字参数报错 ---
                val_tuple = (name, ctype, credit, gpa, status)
                self.root.after(0, self._safe_insert, val_tuple)

                if not is_excluded:
                    total_credit += credit
                    total_score += (credit * gpa)

            final_gpa = 0.0
            if total_credit > 0:
                final_gpa = total_score / total_credit

            result_text = f"总学分: {total_credit:.1f}   |   加权平均绩点: {final_gpa:.4f}"
            self.root.after(0, lambda: self.lbl_result.config(text=result_text))
            self.log("✅ 计算完成！")

        except Exception as e:
            raise e


if __name__ == "__main__":
    root = tk.Tk()
    app = GradeCrawlerGUI(root)
    root.mainloop()
