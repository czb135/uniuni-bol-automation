from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# 启动浏览器
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# 打开你的表单
print("正在打开网页，请稍等...")
driver.get("https://app.smartsheet.com/b/form/a2a520ba7d614e88a00d211941d13364")

# 等待网页彻底加载完（给你10秒钟）
time.sleep(10)

# 获取网页现在的 HTML 代码
html_content = driver.page_source

# 保存到电脑上
with open("smartsheet_code.txt", "w", encoding="utf-8") as f:
    f.write(html_content)

print("✅ 成功！")
print("请在你的文件夹里找到 'smartsheet_code.txt' 这个文件，")
print("把里面的内容（或者文件本身）发给 Gemini。")

driver.quit()