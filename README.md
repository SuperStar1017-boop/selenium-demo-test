# Selenium+Excel数据驱动自动化测试项目
## 项目简介
基于Python + Selenium实现Web自动化测试，使用Excel管理测试用例，读取用例执行测试，自动生成测试报告。
## 技术栈
- Python
- Selenium
- openpyxl（Excel读写）
- Edge浏览器驱动
## 项目文件说明
- `execute_from_excel.py`：主执行脚本，读取Excel用例，执行自动化操作
- `generate_excel.py`：生成测试用例Excel模板
- `generate_full_cases.py`：批量生成完整测试用例
- `generate_report.py`：生成HTML测试报告
- `*.xlsx`：测试用例文件
- `msedgedriver.exe`：Edge浏览器驱动
## 运行步骤
1. 安装依赖：`pip install selenium openpyxl`
2. 准备Excel测试用例
3. 运行主脚本：`python execute_from_excel.py`
4. 查看生成的测试报告
## 项目收获
掌握数据驱动测试思想，熟悉Selenium元素定位，Excel读写处理，自动化测试报告生成，适合Web测试实习岗位。
