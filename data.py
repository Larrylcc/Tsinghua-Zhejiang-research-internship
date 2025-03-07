from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

LOGIN_URL = "https://edu.ybc365.com/"
USERNAME = "yourUsername"
PASSWORD = "yourPassword"

driver = webdriver.Chrome()
driver.maximize_window()  # 最大化窗口，避免元素被遮挡
driver.get(LOGIN_URL)

# 等待登录按钮
login_button = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, "//button/span[text()='登录']"))
)

# 登录
username_input = driver.find_element(By.NAME, "username")
password_input = driver.find_element(By.NAME, "password")
username_input.send_keys(USERNAME)
password_input.send_keys(PASSWORD)
login_button.click()

time.sleep(3)

# 获取 Cookies
cookies = driver.get_cookies()
session_cookies = {cookie['name']: cookie['value'] for cookie in cookies}
print("登录成功，Cookies:", session_cookies)

# 点击评估管理按钮
evaluation_button = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[1]/div/ul/div[5]/a/li/span"))
)
evaluation_button.click()
print("成功进入评估管理页面")

# 等待表格加载
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//table"))
)

# 修改定位方式
report_xpath = "//button[contains(@class, 'el-button') and .//span[contains(text(), '测评报告(新)')]]"

# 等待按钮可点击
WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, report_xpath))
)

# 获取所有报告按钮
report_buttons = driver.find_elements(By.XPATH, report_xpath)
print(f"找到 {len(report_buttons)} 份测评报告")

# 处理每个报告
for index in range(len(report_buttons)):
    print(f"正在处理第 {index + 1} 个报告")

    try:
        # 重新获取按钮列表以避免 stale element 问题
        current_buttons = driver.find_elements(By.XPATH, report_xpath)
        button = current_buttons[index]

        # 确保按钮可见和可点击
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
        time.sleep(1)

        # 使用 JavaScript 点击按钮
        driver.execute_script("arguments[0].click();", button)

        # 等待报告页面加载
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//h1[contains(text(),'测评报告')]"))
        )

        print(f"第 {index + 1} 个报告已打开，等待 3 秒")
        time.sleep(1)

        # 返回
        driver.back()

        # 等待评估管理页面重新加载
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//table"))
        )

        print(f"已返回评估管理页面")
        time.sleep(1)  # 添加短暂延迟

    except Exception as e:
        print(f"处理第 {index + 1} 个报告时出错: {str(e)}")
        continue

print("所有报告处理完成！")
#driver.quit()