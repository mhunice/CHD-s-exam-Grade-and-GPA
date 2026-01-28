import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import pandas as pd
import time
import re
from io import StringIO
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class GradeCrawlerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("教务系统绩点计算器 (Edge 免驱动版)")
        self.root.geometry("900x700")

        self.LOGIN_URL = "http://bkjw.chd.edu.cn/eams/home.action"
        self.GRADE_URL = "http://bkjw.chd.edu.cn/eams/teach/grade/course/person.action"
        self.default_excludes = "科学探索与技术创新（2022）|经典阅读与写作沟通|社会科学与公共责任"

        self.create_widgets()
        self.entry_user.insert(0, "2025903500")
        self.entry_pwd.insert(0, "Marksheep77")

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

        # 3. 日志与结果区域
        frame_log = ttk.LabelFrame(self.root, text="运行日志", padding=10)
        frame_log.pack(fill="x", padx=10, pady=5)
        self.text_log = scrolledtext.ScrolledText(frame_log, height=8, state='disabled', font=("Consolas", 9))
        self.text_log.pack(fill="both", expand=True)

        frame_result = ttk.LabelFrame(self.root, text="成绩明细与结果", padding=10)
        frame_result.pack(fill="both", expand=True, padx=10, pady=5)
        self.lbl_result = ttk.Label(frame_result, text="等待计算...", font=("微软雅黑", 12, "bold"), foreground="blue")
        self.lbl_result.pack(pady=5)

        columns = ("课程名称", "课程类别", "学分", "绩点", "状态")
        self.tree = ttk.Treeview(frame_result, columns=columns, show="headings", height=15)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")
        self.tree.column("课程名称", width=250, anchor="w")

        scrollbar = ttk.Scrollbar(frame_result, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def log(self, message):
        self.root.after(0, self._append_log, message)

    def _append_log(self, message):
        self.text_log.config(state='normal')
        self.text_log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.text_log.see(tk.END)
        self.text_log.config(state='disabled')

    def _safe_insert(self, values):
        self.tree.insert("", "end", values=values)

    def start_thread(self):
        username = self.entry_user.get()
        password = self.entry_pwd.get()
        excludes = self.entry_exclude.get()
        if not username or not password:
            messagebox.showwarning("提示", "请输入学号和密码")
            return
        self.btn_start.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.lbl_result.config(text="正在初始化 Edge 浏览器...")
        thread = threading.Thread(target=self.run_crawler, args=(username, password, excludes))
        thread.daemon = True
        thread.start()

    def run_crawler(self, username, password, exclude_pattern):
        driver = None
        try:
            self.log("🚀 正在通过 Edge Manager 准备环境...")

            # 使用 Edge 配置
            edge_options = EdgeOptions()
            edge_options.add_argument('--headless')  # 静默模式
            edge_options.add_argument('--disable-gpu')
            # 解决某些权限导致驱动获取失败的问题
            os.environ['WDM_SSL_VERIFY'] = '0'

            # 直接启动，Selenium 4 会自动寻找 msedgedriver
            driver = webdriver.Edge(options=edge_options)

            self.log("访问长安大学教务系统...")
            driver.get(self.LOGIN_URL)
            wait = WebDriverWait(driver, 15)

            user_el = wait.until(EC.presence_of_element_located((By.ID, "username")))
            pwd_el = driver.find_element(By.ID, "password")

            user_el.clear()
            user_el.send_keys(username)
            driver.execute_script("arguments[0].value = arguments[1];", pwd_el, password)
            pwd_el.send_keys(Keys.ENTER)

            time.sleep(3)
            self.log("正在跳转成绩页...")
            driver.get(self.GRADE_URL)
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "gridtable")))

            self.process_data(driver.page_source, exclude_pattern)

        except Exception as e:
            self.log(f"❌ 运行出错: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("驱动错误",
                                                            "无法启动浏览器驱动。请确保已安装 Edge 浏览器，或尝试更新 Selenium 库。"))
        finally:
            if driver: driver.quit()
            self.root.after(0, lambda: self.btn_start.config(state="normal"))

    def process_data(self, html, exclude_pattern):
        try:
            tables = pd.read_html(StringIO(html))
            df = None
            for t in tables:
                if '课程' in "".join(t.columns.astype(str)):
                    df = t
                    break

            if df is None: return self.log("未找到成绩表格")

            mapping = {}
            for col in df.columns:
                c = str(col)
                if "课程名称" in c:
                    mapping["name"] = col
                elif "类别" in c:
                    mapping["type"] = col
                elif "学分" in c:
                    mapping["credit"] = col
                elif "绩点" in c:
                    mapping["gpa"] = col

            total_credit = 0.0
            total_points = 0.0

            for _, row in df.iterrows():
                try:
                    name = str(row[mapping["name"]])
                    ctype = str(row[mapping["type"]])
                    credit = pd.to_numeric(row[mapping["credit"]], errors='coerce')
                    gpa = pd.to_numeric(row[mapping["gpa"]], errors='coerce')

                    if pd.isna(credit) or pd.isna(gpa): continue

                    is_excluded = False
                    if exclude_pattern and (re.search(exclude_pattern, name) or re.search(exclude_pattern, ctype)):
                        is_excluded = True

                    status = "排除" if is_excluded else "计入"
                    self.root.after(0, self._safe_insert, (name, ctype, credit, gpa, status))

                    if not is_excluded:
                        total_credit += credit
                        total_points += (credit * gpa)
                except:
                    continue

            final_gpa = total_points / total_credit if total_credit > 0 else 0
            res = f"计入学分: {total_credit:.1f}  |  加权平均绩点: {final_gpa:.4f}"
            self.root.after(0, lambda: self.lbl_result.config(text=res))
            self.log("✅ 成功！")
        except Exception as e:
            self.log(f"解析失败: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GradeCrawlerGUI(root)
    root.mainloop()
