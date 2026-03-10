# -*- coding: utf-8 -*-
from lxml import etree as ET  # 必须使用 lxml 才能用 .getparent()
import json
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QLineEdit, QLabel, QMessageBox, QHBoxLayout, QApplication
)
import sys

from read_data import decrypt_data



def read_file(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    a_data = {
        '项目信息': dict(root.attrib),
        '招标信息': dict(root.find("招标信息").attrib) if root.find("招标信息") is not None else {},
        '投标信息': dict(root.find("投标信息").attrib) if root.find("投标信息") is not None else {},
        '招标控制价信息': dict(root.find("招标控制价信息").attrib) if root.find("招标控制价信息") is not None else {},
        '单项工程': []
    }

    # 解密
    bid = a_data['投标信息']
    if "加密锁号" in bid:
        bid['加密锁号_解密'] = decrypt_data(bid['加密锁号'], 's')
    if "MAC地址" in bid:
        bid['MAC地址_解密'] = decrypt_data(bid['MAC地址'], 's')

    for danxiang in root.findall("单项工程"):
        dx = dict(danxiang.attrib)
        dx['单位工程'] = []
        a_data['单项工程'].append(dx)

        for danwei in danxiang.findall("单位工程"):
            dw = dict(danwei.attrib)
            dw['分部清单'] = []
            dw['措施清单'] = []
            dx['单位工程'].append(dw)

            # 分部清单
            for qd in danwei.findall(".//清单"):
                q = qd.attrib
                item = {
                    "编码": q.get('编码'),
                    "名称": q.get('名称'),
                    "项目特征": q.get('项目特征'),
                    "单位": q.get('单位'),
                    "数量": q.get('数量'),
                    "综合单价": q.get('综合单价'),
                    "综合合价": q.get('综合合价'),
                }
                # 设备费
                for fx in qd.findall(".//单价分析费用项"):
                    if fx.get("名称") == "设备费":
                        jine = float(fx.get("金额") or 0)
                        sl = float(q.get("数量") or 0)
                        dj = round(jine / sl, 2) if sl else 0
                        item['设备单价'] = dj
                        item['综合单价_含设备'] = float(item["综合单价"] or 0) + dj
                dw['分部清单'].append(item)

            # 措施项目（跳过总价措施）
            for cs in danwei.findall(".//措施项目计价表"):
                parent = cs.getparent()
                if parent is not None and parent.get("名称") == "其他总价措施项目":
                    continue
                c = cs.attrib
                dw['措施清单'].append({
                    "编码": c.get('编号'),
                    "名称": c.get('名称'),
                    "项目特征": c.get('项目特征'),
                    "单位": c.get('单位'),
                    "数量": c.get('数量'),
                    "综合单价": c.get('单价'),
                    "综合合价": c.get('合价'),
                })

            # 其他项目
            other = danwei.find("其他项目")
            if other is not None:
                dw['其他项目'] = {}
                for tag, key in [
                    (".//暂列金额明细", '暂列金额明细'),
                    (".//暂估价材料明细", '暂估价材料明细'),
                    (".//暂估价设备明细", '暂估价设备明细'),
                    (".//专业工程暂估明细", '专业工程暂估明细'),
                    (".//计日工项", '计日工项'),
                    (".//总承包服务费项", '总承包服务费项')
                ]:
                    node = other.find(tag)
                    if node is not None:
                        dw['其他项目'][key] = dict(node.attrib)

    return a_data


class XAZBLoader(QWidget):
    def __init__(self):
        super().__init__()  # 传入你的 read_file 函数
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("招标文件加载器")
        self.resize(600, 200)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint)

        layout = QVBoxLayout(self)

        # 1. 文件路径显示
        path_layout = QHBoxLayout()
        self.lbl_path = QLabel("未选择文件")
        self.lbl_path.setWordWrap(True)
        self.lbl_path.setStyleSheet("QLabel { border: 1px solid #aaa; padding: 8px; background: #f8f8f8; }")
        btn_browse = QPushButton("浏览招标文件 (.XAZB)")
        btn_browse.clicked.connect(self.browse_file)
        path_layout.addWidget(self.lbl_path, 1)
        path_layout.addWidget(btn_browse)

        # 2. JSON 保存名输入
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("保存JSON文件名："))
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("例如：xinan_beidi_xiside")
        self.edit_name.setText("output")
        name_layout.addWidget(self.edit_name)

        # 3. 加载按钮
        btn_load = QPushButton("加载并保存为 JSON")
        btn_load.setStyleSheet("QPushButton { background: #4CAF50; color: white; padding: 10px; font-size: 16px; }")
        btn_load.clicked.connect(self.start_convert)

        layout.addLayout(path_layout)
        layout.addLayout(name_layout)
        layout.addWidget(btn_load)
        layout.addStretch()

        self.current_file = ""

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择招标文件",
            "./测试数据/招标数据/",
            "XAZB Files (*.XAZB);;All Files (*)"
        )
        if file_path:
            self.current_file = file_path
            self.lbl_path.setText(file_path)

    def start_convert(self):
        if not self.current_file:
            QMessageBox.warning(self, "错误", "请先选择招标文件！")
            return

        json_name = self.edit_name.text().strip()
        if not json_name:
            QMessageBox.warning(self, "错误", "请输入保存的JSON文件名！")
            return

        try:
            # 读取数据
            self.setEnabled(False)
            btn = self.sender()
            btn.setText("正在加载...")
            data = read_file(self.current_file)

            # 保存 JSON
            default_dir = "./测试数据/招标数据/"
            os.makedirs(default_dir, exist_ok=True)
            json_path = os.path.join(default_dir, f"{json_name}.json")

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            QMessageBox.information(
                self,
                "成功",
                f"招标文件加载完成！\n\n已保存为：\n{json_path}"
            )
            btn.setText("加载并保存为 JSON")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载失败：\n{str(e)}")
            btn.setText("加载并保存为 JSON")
        finally:
            self.setEnabled(True)





if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = XAZBLoader()
    window.show()
    sys.exit(app.exec())