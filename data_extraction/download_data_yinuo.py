#自动打开以诺机构的测评报告页面
#无法自动关闭打开的测评报告页面
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

LOGIN_URL = "https://pre.edu.ybc365.com/"
USERNAME = "yourusername"
PASSWORD = "yourpassword"

# 设置 Chrome 选项
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": r"/Users/larry/Desktop/test",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
}
chrome_options.add_experimental_option("prefs", prefs)

# 使用配置好的选项初始化浏览器
driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()  # 最大化窗口，避免元素被遮挡
driver.get(LOGIN_URL)

# 处理安全警告页面（如果存在）
try:
    # 等待"高级"按钮出现，最多等待5秒
    advanced_button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id=\"details-button\"]"))
    )
    print("检测到安全警告页面，点击'高级'按钮...")
    advanced_button.click()

    # 等待"继续前往"链接出现
    proceed_link = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id=\"proceed-link\"]"))
    )
    print("点击'继续前往'链接...")
    proceed_link.click()

    print("成功绕过安全警告")
except TimeoutException:
    print("未检测到安全警告页面，直接继续登录流程")

# 等待登录页面加载
try:
    # 等待登录按钮
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button/span[text()='登录']"))
    )

    # 登录
    username_input = driver.find_element(By.NAME, "username")
    password_input = driver.find_element(By.NAME, "password")
    username_input.send_keys(USERNAME)
    password_input.send_keys(PASSWORD)
    login_button.click()

    print("已提交登录信息")
except Exception as e:
    print(f"登录过程中出错: {str(e)}")
    driver.save_screenshot("login_error.png")

time.sleep(3)

# 其余代码保持不变
# 获取 Cookies
cookies = driver.get_cookies()
session_cookies = {cookie['name']: cookie['value'] for cookie in cookies}
print("登录成功，Cookies:", session_cookies)

# ... 后续代码不变 ...

# 点击评估管理下拉栏
dropdown_button = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='app']/div/div[1]/div[2]/div[1]/div/ul/div[5]/li/div/i[2]"))
)
dropdown_button.click()
print("成功点击评估管理下拉栏")

# 等待下拉菜单展开
time.sleep(1)

# 点击评估管理菜单项
evaluation_menu_item = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='app']/div/div[1]/div[2]/div[1]/div/ul/div[5]/li/ul/div[1]/a/li"))
)
evaluation_menu_item.click()
print("成功进入评估管理页面")

# 等待表格加载
WebDriverWait(driver, 3).until(
    EC.presence_of_element_located((By.XPATH, "//table"))
)

# 找到并点击下拉菜单，选择50条/页
dropdowns = WebDriverWait(driver, 3).until(
    EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'el-select-dropdown')]"))
)

if len(dropdowns) >= 2:  # 确保有两个元素
    second_dropdown = dropdowns[2]  # 选择第二个元素
    driver.execute_script("arguments[0].click();", second_dropdown)  # 用JS点击，避免不可见问题
else:
    print("未找到足够的下拉菜单元素")

# 查找并点击文本为"50条/页"的选项
option_xpath = "//li/span[contains(text(), '50条/页')]/.."

try:
    fifty_option = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, option_xpath))
    )
    driver.execute_script("arguments[0].click();", fifty_option)  # 用 JS 点击，防止不可见
    print("成功选择 '50条/页'")
except Exception as e:
    print("未找到 '50条/页' 选项:", e)

# 获取所有报告按钮的 XPath - 使用更精确的定位方式
report_xpath = "//button[.//span[normalize-space(text())='测评报告(新)']]"

# 进入 `while` 循环之前，先打印页面上所有按钮的信息
all_buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'el-button')]")
print("\n所有按钮信息：")
for i, btn in enumerate(all_buttons):
    try:
        print(f"\n按钮 {i+1}:")
        print(f"文本内容: {btn.text}")
        print(f"HTML: {btn.get_attribute('outerHTML')}")
        print(f"class属性: {btn.get_attribute('class')}")
        print("-" * 50)
    except Exception as e:
        print(f"获取按钮 {i+1} 信息时出错: {str(e)}")

# 获取符合条件的报告按钮
report_buttons = driver.find_elements(By.XPATH, report_xpath)
print(f"\n找到的测评报告(新)按钮数量: {len(report_buttons)}")

if len(report_buttons) == 0:
    print("没有找到'测评报告(新)'按钮，请检查页面内容")
    driver.save_screenshot("page_buttons.png")
    driver.quit()
    exit()

# 进入 `while` 循环，直到 `btn_next` 按钮不可用
while True:
    # 等待报告按钮加载
    WebDriverWait(driver, 3).until(
        EC.presence_of_all_elements_located((By.XPATH, report_xpath))
    )

    # 获取当前页面所有测评报告按钮
    report_buttons = driver.find_elements(By.XPATH, report_xpath)
    print(f"找到 {len(report_buttons)} 份测评报告")

    # 遍历当前页面的测评报告
    for index in range(50):  # 修改这里，从50开始到99（实际对应第51-100个报告）
        print(f"正在处理第 {index + 1} 个报告")

        try:
            # 重新获取按钮列表以避免 stale element 问题
            current_buttons = driver.find_elements(By.XPATH, report_xpath)
            button = current_buttons[index % 50]  # 修改这里，使用模运算获取当前页面的正确按钮索引

            # 确保按钮可见和可点击
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
            time.sleep(1)

            # 使用 JavaScript 点击按钮
            driver.execute_script("arguments[0].click();", button)

            # 等待报告页面加载，增加等待时间
            WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//h1[contains(text(),'测评报告')]"))
            )
            time.sleep(1)  # 额外等待页面完全加载

            # 等待并点击导出按钮
            try:
                # 使用数据属性和类名组合定位
                export_xpath = "//button[@data-v-2f3849a6 and contains(@class, 'el-button--primary')]//div[contains(text(), '导出')]/.."
                export_button = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, export_xpath))
                )

                # 确保按钮在视图中
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_button)
                time.sleep(2)  # 等待滚动完成

                # 使用 JavaScript 强制点击
                driver.execute_script("""
                    var evt = new MouseEvent('click', {
                        bubbles: true,
                        cancelable: true,
                        view: window
                    });
                    arguments[0].dispatchEvent(evt);
                """, export_button)

                print(f"第 {index + 1} 个报告已导出")
                time.sleep(3)  # 等待下载开始
            except Exception as e:
                print(f"导出按钮操作失败: {str(e)}")
                # 尝试截图记录失败状态
                driver.save_screenshot(f"error_screenshot_{index}.png")
                continue

            # 返回评估管理页面
            driver.back()

            # 等待评估管理页面重新加载
            WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//table"))
            )

            print(f"已返回评估管理页面")
            time.sleep(1)  # 添加短暂延迟

        except Exception as e:
            print(f"处理第 {index + 1} 个报告时出错: {str(e)}")
            continue

    print("当前页面的所有报告处理完成，检查是否有下一页...")

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)  # 等待页面滚动完成
    # 找到 `下一页` 按钮
    try:
        next_button = driver.find_element(By.CLASS_NAME, "btn-next")

        # 检查按钮是否被 `disabled`
        if next_button.get_attribute("disabled") == "disabled":
            print("已到最后一页，结束循环。")
            break  # 退出 `while` 循环

        # 按钮可用，点击进入下一页
        driver.execute_script("arguments[0].click();", next_button)
        print("进入下一页...")
        time.sleep(3)  # 等待页面加载

    except Exception as e:
        print(f"找不到 '下一页' 按钮或出错: {str(e)}")
        break  # 退出循环

print("所有页面的测评报告处理完成！")

# 关闭浏览器
driver.quit()
