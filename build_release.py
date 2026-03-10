import os
import platform
import shutil
import subprocess
import sys

from statics import StaticSource

# 配置
APP_NAME = StaticSource.get_software_name()
MAIN_PY = "main.py"
WIN_ICON = os.path.abspath(os.path.join("icons", "win_icon.ico"))
MAC_ICON = os.path.abspath(os.path.join("icons", "mac_icon.icns"))
PUBLISHER = "雄安空指针"
URL = "https://www.willchalighter.cn"


# 其他方法
def run_cmd(cmd, shell=True):
    print(f"\n>>> {cmd}")
    result = subprocess.run(cmd, shell=shell)
    if result.returncode != 0:
        print(f"✗ 失败！错误码: {result.returncode}")
        sys.exit(1)
    print("✓ 成功")


def build_macos():
    print("\n开始打包 macOS 版本（调用专业 DMG 脚本）")

    # 1. 清理
    for p in ["dist", "build", "build_mac", "__pycache__"]:
        if os.path.exists(p):
            shutil.rmtree(p, ignore_errors=True)

    # 2. PyInstaller 打包
    run_cmd("pyinstaller MyApp.spec --clean")

    if not os.path.exists(f"dist/{APP_NAME}.app"):
        print("打包失败！")
        sys.exit(1)

    # 3. 调用你原来的专业 DMG 脚本
    result = subprocess.run(["./make_dmg.sh"], capture_output=True, text=True)

    if result.returncode != 0:
        print("DMG 打包失败：")
        print(result.stderr)
        sys.exit(1)

    print(result.stdout)
    print("macOS 打包完成！双击 DMG → 拖拽安装 → 秒开无崩溃")


def main():
    version = StaticSource.get_current_version()
    print(f"=" * 60)
    print(f"   雄安清标 v{version} 一键发版神器")
    print(f"   系统: {platform.system()} {platform.machine()}")
    print(f"=" * 60)

    os.makedirs("release", exist_ok=True)

    build_macos()

    print(f"\n发版完成！所有文件已保存到：")
    print(f"   {os.path.abspath('release')}")
    for f in sorted(os.listdir("release")):
        size = os.path.getsize(f"release/{f}") / (1024 * 1024)
        print(f"   📦 {f}  ({size:.1f} MB)")

    print(f"\n直接上传 release 文件夹到 Gitee 即可！用户一键更新！")


if __name__ == "__main__":
    main()
