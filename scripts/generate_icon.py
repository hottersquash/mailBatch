"""Generate MailBatch application icon (mailbatch.ico)."""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QApplication


def draw_icon(size: int) -> QPixmap:
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    margin = max(2, size // 10)
    body = pixmap.rect().adjusted(margin, margin * 2, -margin, -margin)

    painter.setBrush(QColor("#2563EB"))
    painter.setPen(QPen(QColor("#1D4ED8"), max(1, size // 32)))
    painter.drawRoundedRect(body, size // 8, size // 8)

    painter.setPen(QPen(QColor("#FFFFFF"), max(2, size // 16)))
    flap_top = body.top() + body.height() // 3
    painter.drawLine(body.left() + body.width() // 8, flap_top, body.center().x(), body.top() + body.height() // 5)
    painter.drawLine(body.right() - body.width() // 8, flap_top, body.center().x(), body.top() + body.height() // 5)
    painter.drawLine(body.left() + body.width() // 8, flap_top, body.right() - body.width() // 8, flap_top)

    painter.end()
    return pixmap


def main() -> int:
    app = QApplication(sys.argv)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    icon_path = assets_dir / "mailbatch.ico"

    base = draw_icon(256)
    if not base.save(str(icon_path), "ICO"):
        print(f"Failed to save icon: {icon_path}")
        return 1

    png_path = assets_dir / "mailbatch.png"
    base.save(str(png_path), "PNG")
    print(f"Icon saved: {icon_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
