# -*- coding: utf-8 -*-
"""تنظيف عناوين البريد والتحقق من الصيغة قبل الإرسال."""

import re
from typing import Optional, Tuple

# حروف عرض صفرية ومسافات غريبة شائعة عند النسخ من الويب أو Word
_ZW_RE = re.compile(r"[\u200b\u200c\u200d\u200e\u200f\ufeff\u2060\u00ad]")


def normalize_email_address(raw: str) -> str:
    if not raw or not isinstance(raw, str):
        return ""
    s = _ZW_RE.sub("", raw)
    s = s.replace("\u00a0", " ").strip()
    s = s.replace("＠", "@")  # @ ياباني كامل العرض
    s = re.sub(r"\s*@\s*", "@", s)
    return s.strip()


def validate_email_syntax(email: str) -> Tuple[bool, Optional[str]]:
    """التحقق من صيغة العنوان (لا يضمن وجود صندوق فعلي على السيرفر)."""
    if not email or "@" not in email:
        return False, "فارغ أو بدون @"
    try:
        from email_validator import EmailNotValidError, validate_email
    except ImportError:
        if re.match(
            r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$",
            email,
        ):
            return True, None
        return False, "صيغة غير صالحة"
    try:
        validate_email(email, check_deliverability=False)
        return True, None
    except EmailNotValidError as e:
        return False, str(e)


def cell_value_to_email_string(cell_value) -> Optional[str]:
    if cell_value is None:
        return None
    if isinstance(cell_value, bool):
        return None
    if isinstance(cell_value, (int, float)):
        return None
    s = str(cell_value).strip()
    return s if s else None
