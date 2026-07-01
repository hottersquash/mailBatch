"""常用邮箱 SMTP 预设与域名自动识别。"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SmtpPreset:
    preset_id: str
    label: str
    host: str
    port: int
    use_tls: bool
    use_ssl: bool
    domains: tuple[str, ...] = ()


PRESET_AUTO = "auto"
PRESET_CUSTOM = "custom"

SMTP_PRESETS: tuple[SmtpPreset, ...] = (
    SmtpPreset(
        preset_id="qq",
        label="QQ 邮箱 / Foxmail",
        host="smtp.qq.com",
        port=465,
        use_tls=False,
        use_ssl=True,
        domains=("qq.com", "foxmail.com"),
    ),
    SmtpPreset(
        preset_id="163",
        label="163 邮箱",
        host="smtp.163.com",
        port=465,
        use_tls=False,
        use_ssl=True,
        domains=("163.com",),
    ),
    SmtpPreset(
        preset_id="126",
        label="126 邮箱",
        host="smtp.126.com",
        port=465,
        use_tls=False,
        use_ssl=True,
        domains=("126.com",),
    ),
    SmtpPreset(
        preset_id="gmail",
        label="Gmail",
        host="smtp.gmail.com",
        port=587,
        use_tls=True,
        use_ssl=False,
        domains=("gmail.com", "googlemail.com"),
    ),
    SmtpPreset(
        preset_id="outlook",
        label="Outlook / Hotmail",
        host="smtp-mail.outlook.com",
        port=587,
        use_tls=True,
        use_ssl=False,
        domains=("outlook.com", "hotmail.com", "live.com", "msn.com"),
    ),
    SmtpPreset(
        preset_id="office365",
        label="Microsoft 365",
        host="smtp.office365.com",
        port=587,
        use_tls=True,
        use_ssl=False,
        domains=(),  # 企业域名无法枚举，仅手动选择
    ),
    SmtpPreset(
        preset_id="icloud",
        label="Apple iCloud",
        host="smtp.mail.me.com",
        port=587,
        use_tls=True,
        use_ssl=False,
        domains=("icloud.com", "me.com", "mac.com"),
    ),
    SmtpPreset(
        preset_id="exmail",
        label="腾讯企业邮箱",
        host="smtp.exmail.qq.com",
        port=465,
        use_tls=False,
        use_ssl=True,
        domains=(),  # 企业域名各异
    ),
)

_PRESET_BY_ID = {preset.preset_id: preset for preset in SMTP_PRESETS}


def extract_email_domain(email: str) -> str:
    text = email.strip().lower()
    if "@" not in text:
        return ""
    return text.rsplit("@", 1)[-1].strip()


def find_preset_by_domain(email: str) -> SmtpPreset | None:
    domain = extract_email_domain(email)
    if not domain:
        return None
    for preset in SMTP_PRESETS:
        if domain in preset.domains:
            return preset
    return None


def get_preset(preset_id: str) -> SmtpPreset | None:
    return _PRESET_BY_ID.get(preset_id)


def preset_combo_items() -> list[tuple[str, str]]:
    """返回 (preset_id, label) 列表，含自动识别与自定义。"""
    items = [(PRESET_AUTO, "自动识别（按邮箱域名）"), (PRESET_CUSTOM, "自定义")]
    items.extend((preset.preset_id, preset.label) for preset in SMTP_PRESETS)
    return items


def encryption_label(use_ssl: bool, use_tls: bool) -> str:
    if use_ssl:
        return "SSL (465)"
    if use_tls:
        return "STARTTLS (587)"
    return "无"


def parse_encryption_label(label: str) -> tuple[bool, bool]:
    if label.startswith("SSL"):
        return False, True
    if label.startswith("STARTTLS"):
        return True, False
    return False, False
