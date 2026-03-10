import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QIcon
from PySide6.QtWidgets import QApplication, QDialog, QVBoxLayout, QLabel, QHBoxLayout

from app_paths import resource_path
from statics import StaticSource
from utils import functions


class AboutPage(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"关于")
        self.setWindowIcon(QIcon(resource_path("icons/about.png")))
        self.software_version = StaticSource.get_current_version()
        self.platform = functions.get_platform()
        self.show_msg = ""
        if self.platform == "mac":
            self.show_msg = "MacOS 版本"
        elif self.platform == "win":
            self.show_msg = "Windows 版本"
        else:
            self.show_msg = "Linux 版本"
        self.init()
        self.setFixedSize(400, 150)

    def init(self):
        middlelayout = QHBoxLayout()
        middlelayout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_leble = QLabel()
        self.icon_leble.setPixmap(QPixmap(resource_path("icons/mac_icon.png")))
        self.icon_leble.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_leble.setFixedSize(50, 50)
        self.icon_leble.setScaledContents(True)
        self.version_label = QLabel(f"版本：{self.show_msg}（2025年10月25日） {self.software_version}")
        self.copyright_label = QLabel("Copyright © 2025 河北雄安空指针电商工作室（个人独资）")
        self.layout = QVBoxLayout()
        middlelayout.addWidget(self.icon_leble)
        self.layout.addLayout(middlelayout)
        self.layout.addWidget(self.version_label)
        self.layout.addWidget(self.copyright_label)

        self.layout.addStretch()
        self.setLayout(self.layout)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AboutPage()
    window.show()
    sys.exit(app.exec())
