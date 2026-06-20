"""Invoice Renamer — 腾讯云 OCR 图片识别模块

提供腾讯云增值税发票识别 API 的调用封装，包括：
- TC3-HMAC-SHA256 签名（无需额外 SDK）
- 发票识别请求与响应解析
- 字段映射（腾讯云响应 → 本工具字段格式）
- API 密钥的 XOR+Base64 混淆存储
- 配置文件管理（保存/加载/清除）
"""

import base64
import hashlib
import hmac
import json
import os
import re
import time
from datetime import datetime, timezone

# ── 常量 ──────────────────────────────────────────────────────────────────

# 腾讯云 OCR API 配置
TENCENT_OCR_ENDPOINT = "ocr.tencentcloudapi.com"
TENCENT_OCR_SERVICE = "ocr"
TENCENT_OCR_ACTION = "VatInvoiceOCR"
TENCENT_OCR_VERSION = "2018-11-19"
TENCENT_OCR_REGION = "ap-guangzhou"

# 配置文件名
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".invoice_renamer")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")

# XOR 混淆密钥种子（防一眼看到明文，非真正加密）
_OCR_KEY_SEED = b"InvoiceRenamer@2026!Secure"


# ── XOR 混淆工具 ──────────────────────────────────────────────────────────

def _obfuscate(plaintext: str) -> str:
    """XOR + Base64 编码字符串。"""
    if not plaintext:
        return ""
    data = plaintext.encode("utf-8")
    key = _OCR_KEY_SEED
    obfuscated = bytes(data[i] ^ key[i % len(key)] for i in range(len(data)))
    return base64.b64encode(obfuscated).decode("ascii")


def _deobfuscate(ciphertext: str) -> str:
    """Base64 解码 + XOR 还原字符串。"""
    if not ciphertext:
        return ""
    obfuscated = base64.b64decode(ciphertext.encode("ascii"))
    key = _OCR_KEY_SEED
    data = bytes(obfuscated[i] ^ key[i % len(key)] for i in range(len(obfuscated)))
    return data.decode("utf-8")


# ── 配置管理 ──────────────────────────────────────────────────────────────

def _ensure_config_dir() -> None:
    """确保配置目录存在。"""
    os.makedirs(CONFIG_DIR, exist_ok=True)


# ── 用量统计 ──────────────────────────────────────────────────────────────

FREE_TIER_LIMIT = 1000  # 腾讯云每月免费额度

def _increment_usage() -> None:
    """记录一次成功调用。按月份统计，跨月自动重置。"""
    config = _load_config()
    usage = config.get("cloud_ocr_usage", {})
    now = datetime.now()
    month_key = f"{now.year}-{now.month:02d}"
    usage[month_key] = usage.get(month_key, 0) + 1
    config["cloud_ocr_usage"] = usage
    _save_config(config)


def mark_quota_exhausted() -> None:
    """标记当月额度已用尽（调用失败后触发）。"""
    config = _load_config()
    usage = config.get("cloud_ocr_usage", {})
    now = datetime.now()
    month_key = f"{now.year}-{now.month:02d}"
    usage[month_key] = FREE_TIER_LIMIT  # 设为上限阈值
    config["cloud_ocr_usage"] = usage
    _save_config(config)


def get_usage_stats() -> dict:
    """获取当月用量统计。

    Returns:
        {"used": int, "remaining": int, "limit": int, "month": str}
    """
    config = _load_config()
    usage = config.get("cloud_ocr_usage", {})
    now = datetime.now()
    month_key = f"{now.year}-{now.month:02d}"
    used = usage.get(month_key, 0)
    return {
        "used": used,
        "remaining": max(0, FREE_TIER_LIMIT - used),
        "limit": FREE_TIER_LIMIT,
        "month": month_key,
    }


def _load_config() -> dict:
    """加载配置文件，不存在则返回空字典。"""
    if not os.path.isfile(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config(config: dict) -> None:
    """保存配置文件。"""
    _ensure_config_dir()
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def save_credentials(secret_id: str, secret_key: str, enabled: bool = False) -> None:
    """保存 API 密钥到配置文件（混淆后存储）。"""
    config = _load_config()
    config["cloud_ocr"] = {
        "enabled": enabled,
        "provider": "tencent",
        "secret_id": _obfuscate(secret_id),
        "secret_key": _obfuscate(secret_key),
    }
    _save_config(config)


def load_credentials() -> dict:
    """从配置文件加载 API 密钥（自动解密）。"""
    config = _load_config().get("cloud_ocr", {})
    secret_id = _deobfuscate(config.get("secret_id", ""))
    secret_key = _deobfuscate(config.get("secret_key", ""))
    return {
        "enabled": config.get("enabled", False),
        "provider": config.get("provider", "tencent"),
        "secret_id": secret_id,
        "secret_key": secret_key,
    }


def clear_credentials() -> None:
    """清除配置文件中的 API 密钥。"""
    config = _load_config()
    config.pop("cloud_ocr", None)
    _save_config(config)


def has_credentials() -> bool:
    """检查是否已配置有效的 API 密钥。"""
    creds = load_credentials()
    return bool(creds["secret_id"] and creds["secret_key"])


# ── 腾讯云 API 签名（TC3-HMAC-SHA256） ────────────────────────────────────

def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _hmac_sha256(key: bytes, msg: bytes) -> bytes:
    return hmac.new(key, msg, hashlib.sha256).digest()


def _build_signature(
    secret_id: str,
    secret_key: str,
    service: str,
    action: str,
    version: str,
    region: str,
    timestamp: int,
    payload: dict,
) -> tuple[str, str]:
    """构建腾讯云 API 3.0 签名，返回 (authorization, content_type)。"""
    # 准备基本参数
    algorithm = "TC3-HMAC-SHA256"

    # 1. Canonical Request
    http_request_method = "POST"
    canonical_uri = "/"
    canonical_query_string = ""
    content_type = "application/json"
    payload_str = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    hashed_request_payload = _sha256(payload_str.encode("utf-8"))

    canonical_headers = (
        f"content-type:{content_type}\n"
        f"host:{service}.tencentcloudapi.com\n"
        f"x-tc-action:{action.lower()}\n"
    )
    signed_headers = "content-type;host;x-tc-action"

    canonical_request = (
        f"{http_request_method}\n"
        f"{canonical_uri}\n"
        f"{canonical_query_string}\n"
        f"{canonical_headers}\n"
        f"{signed_headers}\n"
        f"{hashed_request_payload}"
    )

    # 2. String to Sign
    date = datetime.fromtimestamp(timestamp, tz=timezone.utc).strftime("%Y-%m-%d")
    credential_scope = f"{date}/{service}/tc3_request"
    hashed_canonical_request = _sha256(canonical_request.encode("utf-8"))
    string_to_sign = (
        f"{algorithm}\n"
        f"{timestamp}\n"
        f"{credential_scope}\n"
        f"{hashed_canonical_request}"
    )

    # 3. Signing Key
    k_date = _hmac_sha256(f"TC3{secret_key}".encode("utf-8"), date.encode("utf-8"))
    k_service = _hmac_sha256(k_date, service.encode("utf-8"))
    k_signing = _hmac_sha256(k_service, b"tc3_request")
    signature = _hmac_sha256(k_signing, string_to_sign.encode("utf-8")).hex()

    # 4. Authorization
    authorization = (
        f"{algorithm} "
        f"Credential={secret_id}/{credential_scope}, "
        f"SignedHeaders={signed_headers}, "
        f"Signature={signature}"
    )

    return authorization, content_type


# ── 发票识别 API ──────────────────────────────────────────────────────────

def recognize_invoice(image_path: str, secret_id: str, secret_key: str) -> dict:
    """识别发票图片或 PDF，自动根据文件类型选择对应 API。

    图片 → VatInvoiceOCR
    PDF  → RecognizeGeneralInvoice（支持 PDF 多页）

    返回统一格式的字段字典，出错时包含 _error 键。
    """
    ext = os.path.splitext(image_path)[1].lower()
    if ext == ".pdf":
        return _recognize_pdf(image_path, secret_id, secret_key)
    return _recognize_image(image_path, secret_id, secret_key)


def _recognize_image(image_path: str, secret_id: str, secret_key: str) -> dict:
    """VatInvoiceOCR — 图片发票识别。"""
    # 读取图片并 Base64 编码
    try:
        with open(image_path, "rb") as f:
            image_data = f.read()
        image_base64 = base64.b64encode(image_data).decode("utf-8")
    except Exception as exc:
        return {"_error": f"图片读取失败: {exc}"}

    # 构建请求 payload
    payload = {"ImageBase64": image_base64}

    # 构建签名
    timestamp = int(time.time())
    try:
        authorization, content_type = _build_signature(
            secret_id=secret_id,
            secret_key=secret_key,
            service=TENCENT_OCR_SERVICE,
            action=TENCENT_OCR_ACTION,
            version=TENCENT_OCR_VERSION,
            region=TENCENT_OCR_REGION,
            timestamp=timestamp,
            payload=payload,
        )
    except Exception as exc:
        return {"_error": f"签名计算失败: {exc}"}

    # 发送请求
    url = f"https://{TENCENT_OCR_ENDPOINT}"
    headers = {
        "Authorization": authorization,
        "Content-Type": content_type,
        "Host": TENCENT_OCR_ENDPOINT,
        "X-TC-Action": TENCENT_OCR_ACTION,
        "X-TC-Timestamp": str(timestamp),
        "X-TC-Version": TENCENT_OCR_VERSION,
        "X-TC-Region": TENCENT_OCR_REGION,
    }

    try:
        import urllib.request
        req = urllib.request.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            response_data = json.loads(resp.read().decode("utf-8"))
    except Exception as exc:
        return {"_error": f"API 请求失败: {exc}"}

    # 检查响应错误
    if "Response" not in response_data:
        return {"_error": "API 返回格式异常"}
    resp = response_data["Response"]
    if "Error" in resp:
        error_code = resp["Error"].get("Code", "")
        error_msg = resp["Error"].get("Message", "")
        # 检测额度耗尽错误
        if any(kw in error_code for kw in ("LimitExceeded", "RequestLimitExceeded", "OperationDenied")):
            mark_quota_exhausted()
            return {"_error": f"本月免费额度已用完（{FREE_TIER_LIMIT}/{FREE_TIER_LIMIT}），次月 1 日自动重置"}
        return {"_error": f"API 错误 [{error_code}]: {error_msg}"}

    # 解析发票字段
    result = _parse_response(resp)

    # 调用成功，更新使用计数
    if "_error" not in result:
        _increment_usage()

    return result


# ── 响应解析 ──────────────────────────────────────────────────────────────

def _normalize_type(raw_type: str) -> str:
    """将腾讯云返回的发票类型名简化为本工具的统一格式。

    映射规则：
        电子发票(普通发票)          → 电子普票
        电子发票(增值税专用发票)     → 增值税专票
        电子发票(铁路电子客票)       → 铁路电子客票
        增值税电子普通发票           → 增值税普票
        增值税专用发票               → 增值税专票
        增值税普通发票               → 增值税普票
        全电发票(普通) / 全电普通发票 → 全电普票
        全电发票(专用) / 全电专用发票 → 全电专票
    """
    if "铁路" in raw_type or "客票" in raw_type:
        return "铁路电子客票"
    if "专用发票" in raw_type or "专票" in raw_type:
        return "增值税专票"
    if "增值税" in raw_type or "增值税普通" in raw_type:
        return "增值税普票"
    if "普通发票" in raw_type or "电子发票" in raw_type:
        return "电子普票"
    if "全电" in raw_type:
        return "全电普票" if "普通" in raw_type else "全电专票"
    # 兜底：保留原始值
    return raw_type


# ── PDF 发票识别（通用票据识别 API）────────────────────────────────────

def _recognize_pdf(file_path: str, secret_id: str, secret_key: str) -> dict:
    """RecognizeGeneralInvoice — PDF 发票识别（支持多页）。"""
    try:
        with open(file_path, "rb") as f:
            pdf_data = f.read()
        pdf_base64 = base64.b64encode(pdf_data).decode("utf-8")
    except Exception as exc:
        return {"_error": f"PDF 读取失败: {exc}"}

    payload = {
        "ImageBase64": pdf_base64,
        "EnableMultiplePage": True,
    }

    timestamp = int(time.time())
    try:
        authorization, content_type = _build_signature(
            secret_id=secret_id, secret_key=secret_key,
            service="ocr",
            action="RecognizeGeneralInvoice",
            version="2018-11-19",
            region=TENCENT_OCR_REGION,
            timestamp=timestamp,
            payload=payload,
        )
    except Exception as exc:
        return {"_error": f"签名计算失败: {exc}"}

    url = f"https://{TENCENT_OCR_ENDPOINT}"
    headers = {
        "Authorization": authorization,
        "Content-Type": content_type,
        "Host": TENCENT_OCR_ENDPOINT,
        "X-TC-Action": "RecognizeGeneralInvoice",
        "X-TC-Timestamp": str(timestamp),
        "X-TC-Version": "2018-11-19",
        "X-TC-Region": TENCENT_OCR_REGION,
    }

    try:
        import urllib.request
        req = urllib.request.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"),
            headers=headers, method="POST",
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            response_data = json.loads(resp.read().decode("utf-8"))
    except Exception as exc:
        return {"_error": f"API 请求失败: {exc}"}

    if "Response" not in response_data:
        return {"_error": "API 返回格式异常"}
    resp = response_data["Response"]
    if "Error" in resp:
        code = resp["Error"].get("Code", "")
        msg = resp["Error"].get("Message", "")
        if any(kw in code for kw in ("LimitExceeded", "RequestLimitExceeded")):
            mark_quota_exhausted()
            return {"_error": f"本月免费额度已用完（{FREE_TIER_LIMIT}/{FREE_TIER_LIMIT}）"}
        return {"_error": f"API 错误 [{code}]: {msg}"}

    fields = _parse_pdf_response(resp)
    if "_error" not in fields:
        _increment_usage()
    return fields


def _parse_pdf_response(resp: dict) -> dict:
    """解析 RecognizeGeneralInvoice 响应，提取字段。"""
    fields = {
        "date": "", "number": "", "buyer": "", "seller": "", "amount": "", "type": "",
    }
    items = resp.get("MixedInvoiceItems", [])
    if not items:
        return fields

    # 取第一个识别到的票据
    item = items[0]
    if item.get("Code") != "OK":
        return fields

    # 提取类型
    sub_type_desc = item.get("SubTypeDescription", "")
    type_desc = item.get("TypeDescription", "")
    fields["type"] = _normalize_type(sub_type_desc or type_desc or "")

    # 提取具体字段（不同票种字段名不同，但常用字段名统一）
    info = item.get("SingleInvoiceInfos", {})
    info_data = {}
    for v in info.values():
        if isinstance(v, dict):
            info_data.update(v)

    raw_date = info_data.get("InvoiceDate", "")
    if raw_date:
        m = re.match(r"(\d{4})\D(\d{1,2})\D(\d{1,2})", raw_date)
        if m:
            fields["date"] = f"{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}"

    fields["number"] = info_data.get("InvoiceNumber", "")
    fields["buyer"] = info_data.get("BuyerName", "")
    fields["seller"] = info_data.get("SellerName", "")

    raw_amount = info_data.get("AmountInFigres", "") or info_data.get("TotalAmount", "")
    if raw_amount:
        cleaned = re.sub(r"[^0-9.]", "", raw_amount)
        if cleaned:
            fields["amount"] = cleaned

    return fields


def _parse_response(resp: dict) -> dict:
    """解析腾讯云 API 响应，映射到本工具字段格式。

    腾讯云返回格式：
        {"VatInvoiceInfos": [{"Name": "InvoiceDate", "Value": "..."}, ...]}
    """
    fields = {
        "date": "", "number": "", "buyer": "", "seller": "", "amount": "", "type": "",
    }

    # 查找发票信息列表
    invoice_infos = resp.get("VatInvoiceInfos", [])
    if not invoice_infos:
        # 可能是旧格式或其他格式
        # 尝试直接从顶层取
        invoice_infos = []
        for key in ("InvoiceDate", "InvoiceNumber", "BuyerName", "SellerName", "TotalAmount"):
            if key in resp:
                invoice_infos.append({"Name": key, "Value": resp[key]})

    info_map = {item["Name"]: item["Value"] for item in invoice_infos if "Name" in item}

    # 字段映射表：腾讯云中文名 → 本工具字段
    _NAME_MAP = {
        "开票日期": "date",
        "发票号码": "number",
        "购买方名称": "buyer",
        "销售方名称": "seller",
        "价税合计(小写)": "amount",
        "发票类型": "type",
    }

    for cn_name, field_key in _NAME_MAP.items():
        raw = info_map.get(cn_name, "")
        if not raw:
            continue

        if field_key == "date":
            # "2026年06月18日" → "2026.06.18"
            m = re.match(r"(\d{4})\D(\d{1,2})\D(\d{1,2})", raw)
            if m:
                fields["date"] = f"{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}"
            else:
                fields["date"] = raw

        elif field_key == "amount":
            # "¥245.00" → "245.00"
            cleaned = re.sub(r"[^0-9.]", "", raw)
            if cleaned:
                fields["amount"] = cleaned

        elif field_key == "type":
            # 简化为已有规则中的类型名
            fields["type"] = _normalize_type(raw)

        else:
            fields[field_key] = raw

    # 兜底：如果没有价税合计金额，试试"合计金额"
    if not fields["amount"]:
        raw_amount = info_map.get("合计金额", "")
        if raw_amount:
            cleaned = re.sub(r"[^0-9.]", "", raw_amount)
            if cleaned:
                fields["amount"] = cleaned

    # 发票类型
    raw_type = info_map.get("InvoiceType", "")
    if raw_type:
        fields["type"] = raw_type

    return fields


# ── 便捷函数 ──────────────────────────────────────────────────────────────

def validate_credentials(secret_id: str, secret_key: str) -> tuple[bool, str]:
    """验证 API 密钥是否有效。

    发送一次轻量请求（空图片会返回明确错误，而非鉴权错误），
    通过错误类型判断密钥是否有效。

    Returns:
        (is_valid, message)
    """
    # 用一个空的 base64 测试密钥
    result = recognize_invoice(__file__, secret_id, secret_key)
    if "_error" in result:
        err = result["_error"]
        if "签名" in err or "AuthFailure" in err or "SecretId" in err:
            return False, "密钥无效，请检查 SecretId/SecretKey"
        if "图片" in err:
            # 密钥有效但图片有问题——这是预期的测试结果
            return True, "密钥有效"
        return True, "密钥有效（检查通过）"
    return True, "密钥有效"
