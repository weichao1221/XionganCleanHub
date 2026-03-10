#!/bin/bash
# ==============================================
# 雄安清标 DMG 打包脚本（精简稳定版）
# ==============================================

APP_NAME="雄安清标"
APP_PATH="./dist/${APP_NAME}.app"
OUTPUT_DIR="./release"
VOLUME_NAME="${APP_NAME}"
VERSION=$(python -c "from statics import get_current_version; print(get_current_version())" 2>/dev/null || echo "unknown")
DATE=$(date "+%Y%m%d")
DMG_NAME="xa_qingbiao_mac_apple_silicon.dmg"
TMP_DIR=$(mktemp -d)

echo "开始打包 DMG → $DMG_NAME"

[[ ! -d "$APP_PATH" ]] && { echo "错误：未找到 $APP_PATH"; exit 1; }

mkdir -p "$OUTPUT_DIR"
mkdir -p "$TMP_DIR/source"

echo "复制应用..."
cp -R "$APP_PATH" "$TMP_DIR/source/" || exit 1

echo "创建 Applications 快捷方式..."
ln -s /Applications "$TMP_DIR/source/Applications"

echo "创建压缩 DMG..."
hdiutil create -srcfolder "$TMP_DIR/source" \
    -volname "$VOLUME_NAME" \
    -fs HFS+ \
    -format UDZO \
    -ov \
    "$OUTPUT_DIR/$DMG_NAME" || exit 1

echo "美化布局..."
hdiutil attach "$OUTPUT_DIR/$DMG_NAME" -noautoopen -quiet
sleep 3

osascript <<APPLESCRIPT
tell application "Finder"
    tell disk "$VOLUME_NAME"
        open
        delay 2
        set theWindow to container window
        set current view of theWindow to icon view
        set toolbar visible of theWindow to false
        set statusbar visible of theWindow to false
        set bounds of theWindow to {200, 200, 800, 600}
        set icon_view_options to icon view options of theWindow
        set arrangement of icon_view_options to not arranged
        set icon size of icon_view_options to 96
        set position of item "${APP_NAME}.app" to {480, 240}
        set position of item "Applications" to {180, 240}
        update without registering applications
        delay 1
        close
    end tell
end tell
APPLESCRIPT

hdiutil detach "/Volumes/$VOLUME_NAME" -force > /dev/null

size=$(du -h "$OUTPUT_DIR/$DMG_NAME" | cut -f1)
echo -e "\n打包成功！"
echo "   文件: $DMG_NAME"
echo "   大小: $size"
echo "   路径: $(pwd)/$OUTPUT_DIR/$DMG_NAME"

rm -rf "$TMP_DIR"
