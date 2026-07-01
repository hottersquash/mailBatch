"""各邮箱 SMTP 授权码/密码获取帮助。"""

from __future__ import annotations

from dataclasses import dataclass

from smtp_presets import PRESET_AUTO, PRESET_CUSTOM, extract_email_domain, find_preset_by_domain, get_preset


@dataclass(frozen=True)
class PasswordHelpInfo:
    provider: str
    summary: str
    steps: str
    help_url: str


PASSWORD_HELP: dict[str, PasswordHelpInfo] = {
    "qq": PasswordHelpInfo(
        provider="QQ 邮箱 / Foxmail",
        summary="QQ 邮箱需使用 SMTP 授权码，不能使用 QQ 登录密码。",
        steps=(
            "1. 登录 mail.qq.com\n"
            "2. 设置 -> 账户\n"
            "3. 开启 POP3/SMTP 服务\n"
            "4. 生成授权码并填入本工具"
        ),
        help_url="https://service.mail.qq.com/detail/0/75",
    ),
    "163": PasswordHelpInfo(
        provider="163 邮箱",
        summary="163 邮箱需开启 SMTP 并使用客户端授权码。",
        steps=(
            "1. 登录 mail.163.com\n"
            "2. 设置 -> POP3/SMTP/IMAP\n"
            "3. 开启 SMTP 服务\n"
            "4. 设置客户端授权密码"
        ),
        help_url="https://help.mail.163.com/faqDetail.do?code=dPF7f2bef975e174791932869046f7679301266",
    ),
    "126": PasswordHelpInfo(
        provider="126 邮箱",
        summary="126 邮箱需开启 SMTP 并使用客户端授权码。",
        steps=(
            "1. 登录 mail.126.com\n"
            "2. 设置 -> POP3/SMTP/IMAP\n"
            "3. 开启 SMTP 服务\n"
            "4. 设置客户端授权密码"
        ),
        help_url="https://help.mail.163.com/faqDetail.do?code=dPF7f2bef975e174791932869046f7679301266",
    ),
    "gmail": PasswordHelpInfo(
        provider="Gmail",
        summary="Gmail 需开启两步验证后创建应用专用密码。",
        steps=(
            "1. 打开 Google 账号安全设置\n"
            "2. 启用两步验证\n"
            "3. 创建“应用专用密码”\n"
            "4. 将 16 位密码填入本工具"
        ),
        help_url="https://support.google.com/accounts/answer/185833",
    ),
    "outlook": PasswordHelpInfo(
        provider="Outlook / Hotmail",
        summary="Outlook 个人邮箱通常需使用应用密码或 Microsoft 账户授权。",
        steps=(
            "1. 登录 Microsoft 账户安全页\n"
            "2. 开启两步验证（如尚未开启）\n"
            "3. 创建应用密码\n"
            "4. 将应用密码填入本工具"
        ),
        help_url="https://support.microsoft.com/account-billing/using-app-passwords-with-apps-that-don-t-support-two-step-verification-5896ed9b-4263-e681-128a-a6fbf9bcf645",
    ),
    "office365": PasswordHelpInfo(
        provider="Microsoft 365",
        summary="企业 Microsoft 365 是否允许 SMTP 由管理员决定。",
        steps=(
            "1. 联系 IT 管理员确认 SMTP 已启用\n"
            "2. 使用企业邮箱账号登录\n"
            "3. 如启用 MFA，创建应用密码\n"
            "4. 主机通常为 smtp.office365.com"
        ),
        help_url="https://learn.microsoft.com/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission",
    ),
    "icloud": PasswordHelpInfo(
        provider="Apple iCloud",
        summary="iCloud 邮件需生成“App 专用密码”。",
        steps=(
            "1. 登录 appleid.apple.com\n"
            "2. 登录与安全性 -> App 专用密码\n"
            "3. 生成新密码\n"
            "4. 将密码填入本工具"
        ),
        help_url="https://support.apple.com/zh-cn/102654",
    ),
    "exmail": PasswordHelpInfo(
        provider="腾讯企业邮箱",
        summary="腾讯企业邮箱通常使用邮箱密码或管理员分配的客户端专用密码。",
        steps=(
            "1. 登录 exmail.qq.com\n"
            "2. 设置 -> 客户端设置\n"
            "3. 开启 SMTP 服务\n"
            "4. 使用邮箱密码或授权码"
        ),
        help_url="https://service.exmail.qq.com/detail?pid=28&version=master&tag=agreement&id=56",
    ),
}

DEFAULT_HELP = PasswordHelpInfo(
    provider="通用邮箱",
    summary="多数邮箱需要使用 SMTP 授权码/应用专用密码，而非网页登录密码。",
    steps=(
        "1. 登录邮箱网页版\n"
        "2. 在设置中查找 SMTP / 客户端授权码\n"
        "3. 开启 SMTP 服务\n"
        "4. 生成授权码后填入本工具"
    ),
    help_url="https://www.google.com/search?q=邮箱+SMTP+授权码+教程",
)


def resolve_password_help_key(email: str, preset_id: str | None = None) -> str:
    if preset_id and preset_id not in (PRESET_CUSTOM, PRESET_AUTO):
        return preset_id

    preset = find_preset_by_domain(email)
    if preset is not None:
        return preset.preset_id

    domain = extract_email_domain(email)
    if domain.endswith("qq.com") or domain.endswith("foxmail.com"):
        return "qq"
    if domain.endswith("163.com"):
        return "163"
    if domain.endswith("126.com"):
        return "126"
    return PRESET_CUSTOM


def get_password_help(email: str, preset_id: str | None = None) -> PasswordHelpInfo:
    key = resolve_password_help_key(email, preset_id)
    if key in PASSWORD_HELP:
        return PASSWORD_HELP[key]

    preset = get_preset(key) if key not in (PRESET_CUSTOM, PRESET_AUTO) else None
    if preset and preset.preset_id in PASSWORD_HELP:
        return PASSWORD_HELP[preset.preset_id]

    help_info = DEFAULT_HELP
    domain = extract_email_domain(email)
    if domain:
        return PasswordHelpInfo(
            provider=f"未知域名 ({domain})",
            summary=help_info.summary,
            steps=help_info.steps,
            help_url=f"https://www.google.com/search?q={domain}+SMTP+授权码",
        )
    return help_info


def format_password_help_tooltip(email: str, preset_id: str | None = None) -> str:
    info = get_password_help(email, preset_id)
    return (
        f"{info.provider}\n"
        f"{info.summary}\n\n"
        f"{info.steps}\n\n"
        f"教程链接：\n{info.help_url}\n\n"
        "点击右侧 ? 按钮可在浏览器打开教程"
    )
