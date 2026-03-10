import os
import ssl
import subprocess
import sys
import urllib.request

from PySide6.QtCore import Signal, QThread
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QProgressBar, QPushButton, QVBoxLayout, QHBoxLayout, QDialog, QMessageBox

from app_paths import resource_path
from statics import StaticSource
from utils import functions


class DownloadProgress(QThread):
    download_progress = Signal(int)
    dmg_path = Signal(str)
    error = Signal(str)

    def __init__(self, url, download_dir):
        super().__init__()
        self.url = url
        self.download_dir = download_dir
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        if not self.url:
            self.error.emit("未获取到有效下载地址。")
            return

        try:
            tokens = StaticSource.get_gitee_token()
            headers = {"Authorization": f"token {tokens}"}
            ssl_context = ssl.create_default_context()
            ssl_context.check_hostname = False
            ssl_context.verify_mode = ssl.CERT_NONE

            req = urllib.request.Request(self.url, headers=headers)
            response = urllib.request.urlopen(req, context=ssl_context)

            total_size = int(response.headers.get('content-length', 0))
            filename = os.path.join(self.download_dir, os.path.basename(self.url))

            downloaded = 0
            block_size = 8192

            with open(filename, 'wb') as f:
                while not self._cancelled:
                    chunk = response.read(block_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        progress = int(downloaded * 100 / total_size)
                    else:
                        progress = 0
                    self.download_progress.emit(progress)

            if self._cancelled:
                if os.path.exists(filename):
                    os.remove(filename)
                return

            if total_size <= 0:
                self.download_progress.emit(100)
            self.dmg_path.emit(filename)
        except Exception as exc:
            self.error.emit(str(exc))


class DownloadWindow(QDialog):
    def __init__(self, download_dir):
        super().__init__()
        self.setWindowTitle("更新进度")
        self.setWindowIcon(QIcon(resource_path("icons/download.png")))
        self.init()
        self.download_dir = download_dir
        self.start_download()

    def init(self):
        layout = QVBoxLayout()
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("开始下载")
        self.start_button.clicked.connect(self.start_download)
        self.cancel_button = QPushButton("取消下载")
        self.cancel_button.clicked.connect(self.cancel_download)
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.cancel_button)

        # ✅ 将进度条设为实例变量
        self.download_bar = QProgressBar()
        self.download_bar.setRange(0, 100)
        self.download_bar.setValue(0)
        self.download_bar.valueChanged.connect(self.update_progress_bar)
        self.download_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #555;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: green;
                width: 10px;
                margin: 0.5px;
            }
        """)

        layout.addWidget(self.download_bar)
        layout.addLayout(button_layout)
        self.setLayout(layout)

    def start_download(self):
        if hasattr(self, 'thread') and self.thread.isRunning():
            return

        url = functions.get_new_version_download_url()
        if not url:
            QMessageBox.warning(self, "更新失败", "暂时无法获取下载地址，请稍后重试。")
            return

        self.cancel_button.setEnabled(True)
        self.thread = DownloadProgress(url, self.download_dir)
        self.thread.download_progress.connect(self.download_bar.setValue)
        self.thread.dmg_path.connect(self.get_filename)
        self.thread.error.connect(self.handle_download_error)
        self.thread.start()
        self.start_button.setEnabled(False)

    def cancel_download(self):
        if hasattr(self, 'thread') and self.thread.isRunning():
            self.thread.cancel()
            self.thread.wait(2000)
            self.download_bar.setValue(0)
            self.start_button.setEnabled(True)
            self.cancel_button.setEnabled(False)

    def get_filename(self, filename):
        self.filename = filename

    def update_progress_bar(self):
        if self.download_bar.value() == 100:
            self.cancel_button.setEnabled(False)
            self.start_button.setEnabled(True)

            # 弹出提示框
            reply = QMessageBox.question(
                self,
                "更新完成",
                "下载已完成，是否立即重启以应用更新？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                system_name = functions.get_platform()
                if system_name == "mac":
                    self.restart_application_macos()
                elif system_name == "win":
                    self.restart_application_windows()

    def handle_download_error(self, error_message):
        self.start_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        QMessageBox.critical(self, "下载失败", error_message)

    def restart_application_windows(self):
        """
        Windows平台下的重启逻辑：
        直接运行下载的exe安装文件，让用户手动完成安装
        """
        # 直接运行下载的exe安装文件
        subprocess.Popen([self.filename], shell=True)

        # 退出当前程序
        QApplication.quit()

    def restart_application_macos(self):
        sh_text = f"""
        #!/bin/bash
        
        DMG_PATH="{self.filename}"
        APP_NAME="雄安清标.app"
        TARGET_DIR="/Applications"
        
        DISK_INFO=$(hdiutil attach "$DMG_PATH")
        DISK_ID=$(echo "$DISK_INFO" | grep "/dev/disk" | awk '{{print $1}}')
        MOUNT_DIR=$(echo "$DISK_INFO" | grep "/Volumes/" | awk '{{print $3}}')
        
        APP_SOURCE=$(find "$MOUNT_DIR" -name "$APP_NAME" -type d | head -n 1)
        if [ -z "$APP_SOURCE" ]; then
            hdiutil detach "$DISK_ID"
            exit 1
        fi
        
        if [ -d "$TARGET_DIR/$APP_NAME" ]; then
            rm -rf "$TARGET_DIR/$APP_NAME"
        fi
        
        cp -R "$APP_SOURCE" "$TARGET_DIR/"
        hdiutil detach "$DISK_ID" -force
        open "$TARGET_DIR/$APP_NAME"
        rm -- "$0"
        """

        script_path = os.path.join(os.path.dirname(self.filename), "update.sh")

        with open(script_path, "w") as f:
            f.write(sh_text)

        os.chmod(script_path, 0o755)  # 设置可执行权限

        # 可选：执行脚本
        # subprocess.Popen(["/bin/bash", script_path])
        subprocess.Popen(["nohup", "bash", script_path, "&"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        # subprocess.Popen([
        #     "open", "-a", "Terminal", script_path
        # ])
        QApplication.quit()



if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = DownloadWindow(download_dir=os.path.join(os.path.expanduser("~"), "Downloads"))
    w.show()
    app.exec()
