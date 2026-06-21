"""Invoice Renamer — PDF 发票解析模块（规则从 rules.json 加载）"""

import json
import re
from pathlib import Path

import pdfplumber

# ── 加载规则文件 ─────────────────────────────────────────────────────────────────────
DEFAULT_RULES = {
    "date": [{"pattern": "开票日期\\s*[：:]?\\s*(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日", "groups": [1,2,3], "sep": "."}],
    "number": [{"pattern": "发票号码[：:]\\s*(\\d+)"}],
    "type_keywords": [{"keyword": "增值税专用发票", "type": "增值税专票"}],
    "buyer_patterns": [],
    "seller_patterns": [],
    "amount_patterns": [],
    "rail_keywords": ["电子客票", "统铁"],
}

def _load_rules() -> dict:
    rules_path = Path(__file__).parent / "rules.json"
    if rules_path.exists():
        try:
            with open(rules_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return DEFAULT_RULES

RULES = _load_rules()


# ── 工具函数 ─────────────────────────────────────────────────────────────────────────

# 主后缀（公司组织形式）——遇到第一个即截断，降低文件名长度
_PRIMARY_SUFFIXES = ["股份有限公司", "有限责任公司", "有限公司", "股份公司", "公司"]
# 次级后缀（经营业态/分支机构）——仅当主后缀未找到时使用
_SECONDARY_SUFFIXES = [
    "分公司", "分店", "加油站", "服务区", "酒店", "宾馆", "饭店",
    "餐厅", "超市", "商行", "门市部", "中心", "车站", "站",
]


def sc(name: str) -> str:
    """截断到公司/机构后缀。

    优先匹配主后缀（公司组织形式），取最早出现的位置截断。
    例如 "湖南投资集团股份有限公司绕城公路西南段分公司" → "湖南投资集团股份有限公司"

    只在第一个空格之前查找后缀，避免多个名称拼在一起时跨名称误匹配
    （如 "XX有限公司 YY站出" 不会把 YY 中的"站"当作后缀）。
    """
    first_space = name.find(' ')
    search_area = name[:first_space] if first_space > 0 else name

    # 第一轮：主后缀，取最早出现
    best_start = -1
    best_end = -1
    for suffix in _PRIMARY_SUFFIXES:
        index = search_area.find(suffix)
        if index != -1 and (best_start == -1 or index < best_start):
            best_start = index
            best_end = index + len(suffix)
    # 第二轮：主后缀未命中，查次级后缀
    if best_start == -1:
        for suffix in _SECONDARY_SUFFIXES:
            index = search_area.find(suffix)
            if index != -1 and (best_start == -1 or index < best_start):
                best_start = index
                best_end = index + len(suffix)
    if best_end > 0:
        return search_area[:best_end]
    # 没找到后缀：有空格则返回第一段，无空格则截断到 15 字符
    if first_space > 0:
        return search_area
    return name[:15] if len(name) > 15 else name


def cp(value: str) -> str:
    value = re.sub(r"\s+", " ", value).strip(" ：:")
    value = re.split(
        r"\s+(?:销|售|买|方|统一社会信用代码|纳税人识别号|地\s*址|开户地址|开户行|账号|备\s*注)[:：]?",
        value,
        maxsplit=1,
    )[0].strip(" ：:")
    return value


def vp(value: str) -> bool:
    if not value or value in {"(章)", "（章）", "章"}:
        return False
    compact = re.sub(r"[\s:：]+", "", value)
    if compact in {"售名称", "销名称", "买名称", "购名称", "销售方名称", "购买方名称"}:
        return False
    if len(re.sub(r"\s+", "", value)) < 2:
        return False
    if any(
        kw in value
        for kw in ["发票号码", "开票日期", "机器编号", "中国铁路祝您旅途愉快", "买票请到12306", "下载次数"]
    ):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", value))


def sp(value: str) -> str:
    value = cp(value)
    if not vp(value):
        return ""
    # 包含公司/机构后缀 → 截断到第一个主后缀
    if re.search(r"(?:" + "|".join(_PRIMARY_SUFFIXES + _SECONDARY_SUFFIXES) + ")", value):
        return sc(value)
    return value[:20] if len(value) > 20 else value


def txid(value: str) -> bool:
    return bool(re.match(r"[0-9A-Z]{15,20}", re.sub(r"\s+", "", value)))


def rx(text: str) -> dict:
    if not any(kw in text for kw in RULES.get("rail_keywords", ["电子客票"])):
        return {}
    data = {}
    match = re.search(
        r"\n([^\s\n]{1,12})\s+[A-Z0-9]+\s+([^\s\n]{1,12})\n站\s+站",
        text,
    )
    if match:
        data["route"] = f"{match.group(1)}-{match.group(2)}"
    match = re.search(
        r"\n[0-9A-Z*]{8,}\s+([^\s\n]{2,10})\n电子客票号",
        text,
    )
    if match:
        data["passenger"] = match.group(1).strip()
    return data


# ── 发票判定关键词 ────────────────────────────────────────────────────────────────

# 第一关 · 强关键词：命中任意一个即直接判定为发票（发票类型标题，基本不会出现在非发票文件中）
_INVOICE_STRONG_KEYWORDS = [
    "发票代码", "增值税专用发票", "增值税普通发票",
    "增值税电子专用发票", "增值税电子普通发票",
    "全电发票", "全电普通发票", "全电专用发票",
    "机动车销售统一发票",
    "电子客票", "统铁",          # 铁路客票
]

# 第二关 · 反关键词：出现在标题区域（前 200 字）则直接判定为非发票
#   这些文件虽然引用了发票信息（如通行费汇总单），但本身不是发票
_INVOICE_EXCLUDE_KEYWORDS = [
    "汇总单", "明细单", "对账单", "结算单", "报销单",
    "清单", "台账", "统计表",
]

# 第三关 · 核心字段：发票必备字段，同时出现基本确定是发票
_INVOICE_CORE_KEYWORDS = [
    "发票号码", "开票日期",
]

# 第三关 · 弱关键词：发票常见字段，单个不足以判定，需配合图片数量
_INVOICE_WEAK_KEYWORDS = [
    "价税合计", "税额", "税率",
    "购买方", "销售方", "购买方名称", "销售方名称",
    "纳税人识别号", "统一社会信用代码",
    "增值税", "电子发票",
]


def _check_invoice(text: str, img_count: int) -> bool:
    """根据第一页文本和图片数量判断是否为发票（三道关卡）。"""
    if not text.strip():
        return False  # 扫描件无文字，无法判断

    # 第一关：强关键词命中即确认
    for kw in _INVOICE_STRONG_KEYWORDS:
        if kw in text:
            return True

    # 第二关：反关键词命中即排除（只看标题区域）
    head = text[:200]
    for kw in _INVOICE_EXCLUDE_KEYWORDS:
        if kw in head:
            return False

    # 第三关：核心字段 + 图片综合判断
    if all(kw in text for kw in _INVOICE_CORE_KEYWORDS):
        return True

    if img_count > 0:
        hit_count = sum(1 for kw in _INVOICE_WEAK_KEYWORDS if kw in text)
        if hit_count >= 2:
            return True

    return False


def is_invoice(file_path: str) -> bool:
    """
    快速判断 PDF 是否为发票（三道关卡，从快到慢）。
    只解析第一页，速度比完整解析快约 5-10 倍。
    """
    try:
        with pdfplumber.open(file_path) as pdf:
            page = pdf.pages[0] if pdf.pages else None
            if page is None:
                return False
            text = page.extract_text() or ""
            img_count = len(page.images)
        return _check_invoice(text, img_count)
    except Exception:
        return False


def parse_image_cloud(file_path: str, secret_id: str, secret_key: str) -> dict:
    """云端 OCR 识别图片发票并提取字段。

    Args:
        file_path: 图片文件路径
        secret_id: 腾讯云 API SecretId
        secret_key: 腾讯云 API SecretKey

    Returns:
        {"fields": dict, "error": str, "not_invoice": bool}
    """
    from cloud_ocr import recognize_invoice

    try:
        result = recognize_invoice(file_path, secret_id, secret_key)
    except Exception as exc:
        return {"fields": {}, "error": f"云端识别失败: {exc}", "not_invoice": True}

    if "_error" in result:
        return {"fields": {}, "error": result["_error"], "not_invoice": True}

    fields = result

    # 用 sp() 截断买卖方名称（保持与 PDF 解析一致）
    if fields.get("buyer"):
        fields["buyer"] = sp(fields["buyer"])
    if fields.get("seller"):
        fields["seller"] = sp(fields["seller"])

    # 检查是否解析出任何有效字段
    has_any = any(fields.get(k) for k in ("date", "number", "buyer", "seller", "amount", "type"))
    if not has_any:
        return {"fields": {}, "error": "云端未识别到发票信息，请确认图片清晰且包含发票", "not_invoice": True}

    return {"fields": fields, "error": "", "not_invoice": False}


def parse_invoice(file_path: str) -> dict:
    """解析发票文件，一次打开 PDF 完成判定和字段提取。"""
    try:
        with pdfplumber.open(file_path) as pdf:
            pages = list(pdf.pages)
            if not pages:
                return {"fields": {}, "error": "PDF 无页面", "not_invoice": True}
            # 提取全部文本 + 首页图片数
            text = "\n".join((p.extract_text() or "") for p in pages)
            text = re.sub(r"(\D)\1{2}", r"\1", text)
            img_count = len(pages[0].images)

            if not text:
                return {"fields": {}, "error": "无法提取文本", "not_invoice": True}

            # 发票判定
            if not _check_invoice(text, img_count):
                return {"fields": {}, "error": "非发票文件", "not_invoice": True}

            # 字段提取
            fields = _extract_fields_from_text(text)
            error = fields.pop("_error", "")
            return {"fields": fields, "error": error, "not_invoice": False}
    except Exception as e:
        return {"fields": {}, "error": str(e), "not_invoice": True}


def _extract_fields_from_text(text: str) -> dict:
    """从已提取的文本中解析发票字段（不打开 PDF）。"""
    result = {
        "date": "", "number": "", "buyer": "", "seller": "", "amount": "", "type": ""
    }
    rail = rx(text)
    _try_date(result, text)
    _try_number(result, text)
    _try_type(result, text, rail)
    _try_buyer_seller(result, text, rail)
    _try_amount(result, text)
    return result


# ── 主解析函数 ─────────────────────────────────────────────────────────────────────
def extract_invoice_fields(file_path: str) -> dict:
    result = {
        "date": "", "number": "", "buyer": "", "seller": "", "amount": "", "type": ""
    }

    try:
        with pdfplumber.open(file_path) as pdf:
            text = "\n".join((page.extract_text() or "") for page in pdf.pages)

        text = re.sub(r"(\D)\1{2}", r"\1", text)

        if not text:
            result["_error"] = "扫描件"
            return result

        rail = rx(text)
        _try_date(result, text)
        _try_number(result, text)
        _try_type(result, text, rail)
        _try_buyer_seller(result, text, rail)
        _try_amount(result, text)

        return result

    except Exception as exc:
        result["_error"] = str(exc)
        return result


# ── 子函数（从 RULES 加载正则）─────────────────────────────────────────────────────

def _try_date(result: dict, text: str) -> None:
    for rule in RULES.get("date", []):
        pattern = rule["pattern"]
        match = re.search(pattern, text)
        if not match:
            continue
        groups = rule.get("groups", [1, 2, 3])
        sep = rule.get("sep", ".")
        try:
            parts = [match.group(g) for g in groups]
            # 对月份和日期做零填充
            if len(parts) >= 2:
                parts[1] = f"{int(parts[1]):02d}"
            if len(parts) >= 3:
                parts[2] = f"{int(parts[2]):02d}"
            result["date"] = sep.join(parts)
            return
        except Exception:
            continue


def _try_number(result: dict, text: str) -> None:
    for rule in RULES.get("number", []):
        pattern = rule["pattern"]
        match = re.search(pattern, text)
        if match:
            result["number"] = match.group(1)
            return


def _try_type(result: dict, text: str, rail: dict) -> None:
    if rail:
        result["type"] = "铁路电子客票"
        return
    for rule in RULES.get("type_keywords", []):
        if rule["keyword"] in text:
            result["type"] = rule["type"]
            return


def _try_buyer_seller(result: dict, text: str, rail: dict) -> None:
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]

    # 购买方
    if not result["buyer"]:
        for pattern in RULES.get("buyer_patterns", []):
            match = re.search(pattern, text)
            if match:
                value = sp(match.group(1))
                if value:
                    result["buyer"] = value
                    break

    # 销售方
    if not result["seller"]:
        for pattern in RULES.get("seller_patterns", []):
            match = re.search(pattern, text)
            if match:
                value = sp(match.group(1))
                if value:
                    result["seller"] = value
                    break

    # 同一行匹配买卖双方
    match = re.search(
        r"(?:购|买)\s*名称[\s：:]*([^\n]{2,20}?)\s+(?:销|售)\s*名称[\s：:]*([^\n]{2,40})",
        text,
    )
    if match:
        if not result["buyer"]:
            result["buyer"] = sp(match.group(1))
        if not result["seller"]:
            result["seller"] = sp(match.group(2))

    # 邻近行匹配
    for index, line in enumerate(lines):
        match = re.match(r"名\s*称[:：]\s*(.+)", line)
        if not match:
            continue
        value = sp(match.group(1))
        if not value:
            continue
        next_line = lines[index + 1] if index + 1 < len(lines) else ""
        if not result["buyer"] and re.search(r"[购买]", next_line):
            result["buyer"] = value
        elif not result["seller"] and re.search(r"[销售]", next_line):
            result["seller"] = value

    # 通过纳税人识别号定位主体
    points = []
    for index in range(len(lines) - 1):
        if txid(lines[index + 1]):
            # 一行可能包含多个名称（用空格分隔，全电发票常见格式），分别处理
            for part in lines[index].split():
                value = sp(part)
                if value and value not in points:
                    points.append(value)
    if not result["buyer"] and points:
        result["buyer"] = points[0]
    if not result["seller"]:
        for value in points:
            if value != result["buyer"]:
                result["seller"] = value
                break

    # 铁路电子客票
    if rail:
        subject = "_".join(part for part in [rail.get("passenger", ""), rail.get("route", "")] if part)
        if subject:
            result["seller"] = subject


def _try_amount(result: dict, text: str) -> None:
    for pattern in RULES.get("amount_patterns", []):
        pattern_str = pattern if isinstance(pattern, str) else pattern["pattern"]
        values = re.findall(pattern_str, text)
        if values:
            amount = re.sub(r"[^0-9.]", "", values[-1])
            result["amount"] = amount.replace(",", "")
            break
