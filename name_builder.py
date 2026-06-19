"""Invoice Renamer — 文件名拼接模块"""

import os
import re


def sanitize_part(value: str) -> str:
    """
    清理文件名片段，移除非法字符并规范化空白。

    非法字符：\\ / : * ? " < > |
    替换规则：上述字符 → "-"，连续空白 → 单个空格
    """
    value = re.sub(r'[\\/:*?"<>|]+', "-", str(value).strip())
    value = re.sub(r"\s+", " ", value)
    return value.strip(" ._")


def build_name(
    fields: dict,
    enabled_map: dict,
    field_order: list,
    custom_text: str = "",
) -> str:
    """
    根据字段配置拼接新文件名。

    fields       : 解析得到的发票字段 dict
    enabled_map  : {key: bool} 各字段是否启用
    field_order  : 字段排序列表
    custom_text  : 自定义字段内容

    返回如 "2024.01.15_增值税专票_某某有限公司_1234.56元_发票.pdf"
    """
    parts = []
    for key in field_order:
        if not enabled_map.get(key):
            continue
        raw_value = custom_text if key == "custom" else fields.get(key, "")
        cleaned = sanitize_part(raw_value)
        if not cleaned:
            continue
        # 金额字段追加"元"
        if key == "amount":
            parts.append(f"{cleaned}元")
        else:
            parts.append(cleaned)

    if not parts:
        return "未识别发票.pdf"
    return "_".join(parts) + ".pdf"


def ensure_unique_name(
    directory: str,
    desired_name: str,
    current_name: str | None = None,
) -> str:
    """
    确保目标文件名在目录中唯一，若重名则追加 _1 / _2 ...

    directory    : 目标目录
    desired_name : 期望的文件名
    current_name : 当前文件名（若目标与当前相同则不修改）
    """
    if desired_name == current_name:
        return desired_name

    base, ext = os.path.splitext(desired_name)
    candidate = desired_name
    index = 1
    while os.path.exists(os.path.join(directory, candidate)) and candidate != current_name:
        candidate = f"{base}_{index}{ext}"
        index += 1
    return candidate
