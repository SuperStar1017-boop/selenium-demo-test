from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
import time

# ★ 关键：明确告诉Python，驱动就在当前文件夹
service = Service(executable_path='./msedgedriver.exe')
driver = webdriver.Edge(service=service)

# 打开SauceDemo网站
driver.get("https://www.saucedemo.com/")

# 输入账号密码
driver.find_element(By.ID, "user-name").send_keys("standard_user")
driver.find_element(By.ID, "password").send_keys("secret_sauce")

# 点击登录按钮
driver.find_element(By.ID, "login-button").click()

# 等待2秒，让页面反应一下
time.sleep(2)

# 检查是否登录成功（断言）
assert "Swag Labs" in driver.title
print("✅ 恭喜！你的第一个自动化测试跑通啦！")

# 关闭浏览器
driver.quit()