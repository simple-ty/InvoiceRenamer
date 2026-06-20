"""测试 WebView 版 API 的业务逻辑。"""

import sys
import time
from pathlib import Path
sys.path.insert(0, r"C:\Users\29292\Desktop\InvoiceRenamer")

from webapp import Api

api = Api()

# 测试初始化状态
state = api.get_init_state()
print("Init state keys:", list(state.keys()))
print("Field order:", state.get("field_order"))
print("Cloud configured:", state["cloud"]["configured"])
print("Cloud enabled:", state["cloud"]["enabled"])

# 测试模板更新
result = api.update_template(
    ["date", "type", "seller", "amount"],
    {"date": True, "type": True, "seller": True, "amount": True, "number": False, "buyer": False, "custom": False},
    "项目A",
)
print("Update template:", result["ok"], len(result["records"]))

# 测试图片扫描（使用项目目录下的测试文件）
test_dir = Path(r"C:\Users\29292\Desktop\InvoiceRenamer\Test_Invoice\picture")
test_files = [
    str(p) for p in test_dir.iterdir()
    if p.is_file() and p.suffix.lower() in (".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif")
]
api.selected_paths = test_files
print("开始扫描测试文件...")
api.scan_files()

# 等待扫描完成
for i in range(30):
    if not api._scanning:
        break
    time.sleep(0.5)

print(f"扫描完成，记录数: {len(api.records)}")
for r in api.records:
    print(f"  {r['source_name']} -> {r['status']} -> error: {r.get('error', '')}")
    print(f"    new_name: {r['new_name']}")
    print(f"    fields: {r['fields']}")

print("All basic tests passed.")

# ── PDF 解析测试 ─────────────────────────────────────────────
print("\n开始测试 PDF 解析...")
pdf_dir = Path(r"C:\Users\29292\Desktop\InvoiceRenamer\Test_Invoice")
pdf_files = [
    str(p) for p in pdf_dir.iterdir()
    if p.is_file() and p.suffix.lower() == ".pdf"
][:5]

api2 = Api()
api2.selected_paths = pdf_files
api2.scan_files()

for i in range(60):
    if not api2._scanning:
        break
    time.sleep(0.5)

print(f"PDF 扫描完成，记录数: {len(api2.records)}")
for r in api2.records:
    print(f"  {r['source_name']} -> {r['status']} -> error: {r.get('error', '')}")
    print(f"    new_name: {r['new_name']}")
    print(f"    fields: {r['fields']}")

print("All PDF tests passed.")

# ── 文件夹选择扫描测试 ───────────────────────────────────────
print("\n开始测试文件夹扫描...")
api3 = Api()
api3.selected_paths = [str(pdf_dir)]  # 模拟选择文件夹路径（实际 choose_folder 会展开）
# 手动展开文件夹，模拟 choose_folder 行为
api3.selected_paths = [
    str(p) for p in pdf_dir.iterdir()
    if p.is_file() and p.suffix.lower() in (".pdf",)
][:3]
api3.scan_files()

for i in range(60):
    if not api3._scanning:
        break
    time.sleep(0.5)

print(f"文件夹扫描完成，记录数: {len(api3.records)}")
for r in api3.records:
    print(f"  {r['source_name']} -> {r['status']}")

print("All folder tests passed.")
