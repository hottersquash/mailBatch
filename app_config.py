"""应用配置持久化：自动保存上次配置，支持导入导出。"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

CONFIG_VERSION = 1
DEFAULT_CONFIG_NAME = "last_config.json"


@dataclass
class AppConfig:
    excel_path: str = ""
    template_path: str = ""
    template_text: str = ""
    attachment_paths: list[str] = field(default_factory=list)
    subject: str = ""
    email_column: str = ""
    cc_column: str = ""
    smtp_host: str = ""
    smtp_port: int = 587
    sender_email: str = ""
    smtp_user: str = ""
    smtp_password: str = ""
    mail_from: str = ""
    use_tls: bool = True
    use_ssl: bool = False
    smtp_preset: str = "auto"
    send_interval: float = 1.0
    last_dialog_dir: str = ""

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["version"] = CONFIG_VERSION
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> AppConfig:
        known = {f.name for f in cls.__dataclass_fields__.values()}  # type: ignore[attr-defined]
        filtered = {key: value for key, value in data.items() if key in known}
        config = cls(**filtered)
        if not config.sender_email:
            config.sender_email = config.smtp_user or config.mail_from
        return config


def config_path(base_dir: str | Path, name: str = DEFAULT_CONFIG_NAME) -> Path:
    return Path(base_dir) / name


def save_app_config(config: AppConfig, path: str | Path) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(config.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")


def load_app_config(path: str | Path) -> AppConfig:
    path = Path(path)
    if not path.is_file():
        return AppConfig()
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("配置文件格式无效")
    return AppConfig.from_dict(data)


def export_app_config(config: AppConfig, path: str | Path) -> None:
    save_app_config(config, path)


def import_app_config(path: str | Path) -> AppConfig:
    return load_app_config(path)
