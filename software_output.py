import json
import openpyxl

# 将软件内容保存在 software.json 文件中
with open('software.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

items = data.get("data", {}).get("items", [])

# 创建 Excel 工作簿
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "软件清单"

# 写入表头
headers = [
    "软件全称",
    "软件版本",
    "操作系统",
    "Bundle ID/软件发布者",
    "已安装终端数",
    "已安装用户",
    "安装率"
]
ws.append(headers)

# 写入每行数据
for item in items:
    software_name = item.get("software_name", "")
    versions = ", ".join(item.get("versions", []))  # 多版本用逗号分隔
    os_name = item.get("os", "")
    bundle_or_publisher = item.get("bundle_id", "")
    publisher = item.get("publisher", "")
    if publisher and publisher != bundle_or_publisher:
        bundle_or_publisher += f" / {publisher}"
    installed_devices = item.get("installed_devices", 0)
    installed_users = item.get("installed_users", 0)
    installed_rate = item.get("installed_rate", 0)

    ws.append([
        software_name,
        versions,
        os_name,
        bundle_or_publisher,
        installed_devices,
        installed_users,
        installed_rate
    ])

# 自动调整列宽（简单处理）
for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

# 保存 Excel 文件
output_file = "软件清单.xlsx"
wb.save(output_file)
print(f"✅ 已导出 {output_file}")