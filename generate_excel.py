import json
import os
from openai import OpenAI
import openpyxl  # 需要先安装: pip install openpyxl

# ======================= 配置区 =======================
API_KEY = "sk-16dd502f68d344de86d5d7b2f9d5634b"   # 替换成你的有效 Key
# =====================================================

print("🚀 程序启动...")

client = OpenAI(
    api_key=API_KEY,
    base_url="https://api.deepseek.com",
    timeout=15
)

def build_prompt():
    return """
请为 Swag Labs (https://www.saucedemo.com) 的登录功能生成测试用例。
覆盖以下场景：正确登录、密码错误、用户名不存在、用户名为空、密码为空、两者都为空、锁定用户。

输出JSON数组，每个用例包含以下字段：
- id: 用例编号
- title: 用例标题
- preconditions: 前置条件
- test_steps: 测试步骤（列表）
- test_data: 测试数据（对象）
- expected_result: 预期结果

只输出JSON，不要解释。
"""

def generate_test_cases():
    print("⏳ 正在调用 API...")
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": "你是测试工程师。"},
            {"role": "user", "content": build_prompt()}
        ]
    )
    return response.choices[0].message.content

def save_to_excel(test_cases, filename="测试用例.xlsx"):
    """将测试用例保存为 Excel 文件"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "登录测试用例"

    # 定义表头（按你要求的列顺序）
    headers = ["用例编号", "用例标题", "项目/模块", "优先级", "前置条件", "测试步骤", "测试数据", "预期结果"]
    ws.append(headers)

    # 填充数据
    for tc in test_cases:
        # 处理测试步骤（列表转字符串，用换行分隔）
        steps = tc.get("test_steps", [])
        if isinstance(steps, list):
            steps_str = "\n".join([f"{i+1}. {s}" for i, s in enumerate(steps)])
        else:
            steps_str = str(steps)

        # 处理测试数据（字典转字符串）
        test_data = tc.get("test_data", {})
        if isinstance(test_data, dict):
            test_data_str = "\n".join([f"{k}: {v}" for k, v in test_data.items()])
        else:
            test_data_str = str(test_data)

        row = [
            tc.get("id", ""),
            tc.get("title", ""),
            "登录模块",           # 项目/模块（固定值，可按需修改）
            "高",                 # 优先级（固定为"高"，可按需修改）
            tc.get("preconditions", ""),
            steps_str,
            test_data_str,
            tc.get("expected_result", "")
        ]
        ws.append(row)

    # 调整列宽
    column_widths = [12, 30, 12, 10, 30, 40, 25, 40]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    wb.save(filename)
    print(f"✅ Excel 文件已保存：{os.path.abspath(filename)}")

if __name__ == '__main__':
    print("📡 开始请求...")
    result = generate_test_cases()
    print("✅ 收到回复，正在解析...")

    # 清理 Markdown
    cleaned = result.strip()
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()

    try:
        test_cases = json.loads(cleaned)
        print(f"🎉 成功生成 {len(test_cases)} 条测试用例！")

        # 保存为 Excel
        save_to_excel(test_cases)

        # 同时在终端打印预览
        print("\n" + "="*60)
        print(json.dumps(test_cases, indent=4, ensure_ascii=False))
        print("="*60)

    except json.JSONDecodeError as e:
        print("⚠️ JSON 解析失败，原始内容：")
        print(result)