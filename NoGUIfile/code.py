#控制台输出版本
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from io import StringIO
from selenium.webdriver.common.action_chains import ActionChains
import pyperclip

def has_captcha(driver):
    # 查找所有可能的验证码输入框
    captcha_inputs = driver.find_elements(
        By.XPATH, "//input[contains(@id,'captcha') or contains(@name,'captcha')]"
    )

    for el in captcha_inputs:
        try:
            if el.is_displayed() and el.size["height"] > 0:
                return True
        except:
            pass

    # 查找验证码图片
    captcha_imgs = driver.find_elements(
        By.XPATH, "//img[contains(@src,'captcha')]"
    )

    for img in captcha_imgs:
        try:
            if img.is_displayed() and img.size["height"] > 0:
                return True
        except:
            pass

    return False

def login_auto_or_manual(driver, username, password):
    wait = WebDriverWait(driver, 15)

    # 等待登录页
    user_input = wait.until(EC.presence_of_element_located((By.ID, "username")))
    user_input.clear()
    user_input.send_keys(username)

    # 判断是否启用验证码
    if has_captcha(driver):
        print("👉 检测到验证码，请手动登录")
        input("登录完成后按回车继续...")
        return True

    print("✅ 无验证码，使用【粘贴方式】自动登录")

    pwd_input = wait.until(EC.element_to_be_clickable((By.ID, "password")))

    # 激活密码框
    driver.execute_script("arguments[0].focus();", pwd_input)
    time.sleep(0.2)
    pwd_input.click()
    time.sleep(0.2)

    # 放入剪贴板
    pyperclip.copy(password)
    time.sleep(0.2)

    # Ctrl + V 粘贴
    from selenium.webdriver.common.action_chains import ActionChains

    # 已确保 pwd_input 已 click + focus
    pyperclip.copy(password)
    time.sleep(0.2)

    actions = ActionChains(driver)
    actions.key_down(Keys.CONTROL) \
        .send_keys('v') \
        .key_up(Keys.CONTROL) \
        .perform()

    time.sleep(0.5)

    # 回车登录
    pwd_input.send_keys(Keys.ENTER)

    # 等待跳转
    time.sleep(3)

    if "login" not in driver.current_url.lower():
        print("🎉 自动登录成功")
        return True
    else:
        print("⚠️ 自动登录失败，切换人工")
        input("请手动登录后按回车继续...")
        return True



# ======================
# 1. 基本配置（你本地填写）
# ======================
USERNAME = ""
PASSWORD = ""

LOGIN_URL = "http://bkjw.chd.edu.cn/eams/home.action"
GRADE_URL = "http://bkjw.chd.edu.cn/eams/teach/grade/course/person.action"

# 要排除的课程类别（与你要求完全一致）
EXCLUDE_KEYWORDS = [
        "科学探索与技术创新（2022）",
        "经典阅读与写作沟通",
        "社会科学与公共责任"
    ]

# ======================
# 2. 启动浏览器
# ======================
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 15)

driver.get(LOGIN_URL)

# ======================
# 3. 登录
# ======================
driver.get(LOGIN_URL)
login_auto_or_manual(driver, USERNAME, PASSWORD)



# 等待首页加载完成
time.sleep(3)

# ======================
# 4. 进入成绩查询页面
# ======================
driver.get(GRADE_URL)

# 等待成绩表格出现
wait.until(EC.presence_of_element_located((By.CLASS_NAME, "gridtable")))

# ======================
# 5. 解析 HTML 表格
# ======================
html = driver.page_source

# EAMS 成绩页通常只有一个主表
df = pd.read_html(html)[0]

# 打印列名，首次运行可对照确认
#print("识别到的列名：")
#print(df.columns)

# ======================
# 6. 标准化列名（防止空格/换行）
# ======================
# 查看原始表格结构（只打印一次用来确认）
from io import StringIO

html = driver.page_source
tables = pd.read_html(StringIO(html))

#print(f"页面中共识别到 {len(tables)} 个表格")

grade_df = None

for i, table in enumerate(tables):
    #print(f"\n表格 {i} 前几行：")
    #print(table.head())

    # 成绩表一定同时包含“课程”和“绩点”相关列
    cols = table.columns.astype(str)

    if any("课程" in c for c in cols) and any("绩点" in c for c in cols):
        grade_df = table
        print(f"\n✅ 已识别为成绩表：表格 {i}")
        break

if grade_df is None:
    raise RuntimeError("❌ 未能在页面中找到成绩表")

df = grade_df.copy()
df.columns = df.columns.astype(str)

# 根据你截图，提取需要的列（用“包含关系”，不怕顺序变）
def find_col(keyword):
    for c in df.columns:
        if keyword in c:
            return c
    raise RuntimeError(f"找不到列：{keyword}")

col_name = find_col("课程")
col_type = find_col("课程类别")
col_credit = find_col("学分")
col_gpa = find_col("绩点")

df = df[[col_name, col_type, col_credit, col_gpa]]
df.columns = ["课程名称", "课程类别", "学分", "绩点"]

df["学分"] = pd.to_numeric(df["学分"], errors="coerce")
df["绩点"] = pd.to_numeric(df["绩点"], errors="coerce")

df = df.dropna(subset=["学分", "绩点"])


# ======================
# 7. 过滤不参与计算的课程
# ======================
import re

pattern = "|".join(EXCLUDE_KEYWORDS)

df = df[
    ~df["课程类别"]
    .astype(str)
    .str.contains(pattern, regex=True)
]


# 只保留计算所需列
df = df[["课程名称", "课程类别", "学分", "绩点"]]

# 转为数值（防止字符串）
df["学分"] = pd.to_numeric(df["学分"])
df["绩点"] = pd.to_numeric(df["绩点"])

# ======================
# 8. 加权绩点计算
# ======================
total_score = (df["学分"] * df["绩点"]).sum()
total_credit = df["学分"].sum()

gpa = total_score / total_credit

# ======================
# 9. 输出结果
# ======================
print("\n====== 计算结果 ======")
print(f"参与计算课程数：{len(df)}")
print(f"总学分：{total_credit:.2f}")
print(f"加权绩点：{gpa:.4f}")

print("\n====== 明细 ======")
print(df)

driver.quit()
