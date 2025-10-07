# plugin_api.py
from dataclasses import dataclass
from typing import Callable, Optional, Any
from pathlib import Path
import sys

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QFrame


@dataclass
class ToolSpec:
    """
    A lightweight spec describing a tool that the ATLAS shell can mount.

    Attributes:
        id:     Unique string id (stable across builds).
        name:   Human-readable name for the sidebar.
        factory:Callable returning a QWidget for this tool.
        icon:   Optional QIcon or string path to an icon.
        order:  Integer sort key (lower comes first).
    """
    id: str
    name: str
    factory: Callable[[], QWidget]
    icon: Optional[Any] = None
    order: int = 100


class BaseToolPage(QWidget):
    """
    Convenience QWidget:
      - Title
      - Optional subtitle
      - Divider line
      - self.content (QVBoxLayout) for tool body
    """
    def __init__(self, title: str, subtitle: str = "", parent: Optional[QWidget] = None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        t = QLabel(title)
        t.setStyleSheet("font-size:20px;font-weight:600;")
        root.addWidget(t)

        if subtitle:
            s = QLabel(subtitle)
            s.setStyleSheet("color:#666;")
            root.addWidget(s)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        root.addWidget(line)

        self.content = QVBoxLayout()
        self.content.setSpacing(8)
        root.addLayout(self.content)
        root.addStretch(1)


def runtime_path(*parts: str) -> Path:
    """
    Resolve a data/resource path in dev and in PyInstaller frozen builds.
    Use this in plugins to find things like 'Template.xlsx'.

    Example:
        tpl = runtime_path("Template.xlsx")
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        base = Path(__file__).resolve().parent
    return (base / Path(*parts)).resolve()
