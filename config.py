"""Invoice Renamer — 常量与配置"""

APP_VERSION = "v0.5.9"
APP_AUTHOR = "Simple"
APP_ID = f"InvoiceRenamer.{APP_VERSION}"

RESULT_HINT = "表格中可直接查看识别结果和拟生成文件名，确认无误后再执行重命名。"

# 图片扩展名
IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif")
ALLOWED_EXTENSIONS = (".pdf",) + IMAGE_EXTENSIONS

FIELD_LABELS = {
    "date": "开票日期",
    "type": "发票类型",
    "number": "发票号码",
    "buyer": "购买方",
    "seller": "销售方",
    "amount": "金额",
    "custom": "自定义字段",
}
DEFAULT_FIELD_ORDER = ["date", "type", "number", "buyer", "seller", "amount", "custom"]
DEFAULT_FIELD_ENABLED = {
    "date": True,
    "type": True,
    "number": False,
    "buyer": False,
    "seller": True,
    "amount": True,
    "custom": False,
}

TREE_TAG_COLORS = {
    "success":     "#191919",
    "partial":     "#FA9D3B",
    "error":       "#FA5151",
    "idle":        "#8A8A8A",
    "not_invoice": "#B0B0B0",   # 非发票文件：灰色，视觉上最低优先级
}

STAT_CARD_STYLE = {
    "total":       {"fg": "#FFFFFF", "accent": "#191919", "title": "文件总数"},
    "complete":    {"fg": "#FFFFFF", "accent": "#07C160", "title": "完整识别"},
    "partial":     {"fg": "#FFFFFF", "accent": "#FA9D3B", "title": "部分识别"},
    "failed":      {"fg": "#FFFFFF", "accent": "#FA5151", "title": "未识别/异常"},
    "not_invoice": {"fg": "#FFFFFF", "accent": "#B0B0B0", "title": "非发票"},
}
