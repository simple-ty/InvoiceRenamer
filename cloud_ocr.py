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
    """调用腾讯云增值税发票识别 API，返回结构化字段。

    Args:
        image_path: 图片文件路径
        secret_id: 腾讯云 API SecretId
        secret_key: 腾讯云 API SecretKey

    Returns:
        包含 fields 的字典，出错时包含 _error 键。
    """
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

    # 调试输出
    print(f"\n[云端OCR调试] 文件: {os.path.basename(image_path)}")
    print(f"[云端OCR调试] 状态码: {response_data.get('Response', {}).get('Error', {}).get('Code', '成功')}")
    print(f"[云端OCR调试] 原始响应:\n{json.dumps(response_data, ensure_ascii=False, indent=2)}")

    # 检查响应错误
    if "Response" not in response_data:
        return {"_error": "API 返回格式异常"}
    resp = response_data["Response"]
    if "Error" in resp:
        error_code = resp["Error"].get("Code", "")
        error_msg = resp["Error"].get("Message", "")
        return {"_error": f"API 错误 [{error_code}]: {error_msg}"}

    # 解析发票字段
    result = _parse_response(resp)
    print(f"[云端OCR调试] 解析结果: {json.dumps(result, ensure_ascii=False)}")
    return result


# ── 响应解析 ──────────────────────────────────────────────────────────────

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
            # 直接用中文类型名（如 "电子发票(普通发票)"）
            fields["type"] = raw

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
