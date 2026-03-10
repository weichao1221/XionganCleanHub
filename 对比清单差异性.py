import sys

from PySide6.QtGui import QStandardItemModel, QStandardItem, QFont
from PySide6.QtWidgets import *


class StructureAligner(QWidget):
    def __init__(self, data1, data2, name1="", name2=""):
        super().__init__()
        self.setWindowTitle("数据结构梳理")
        self.resize(600, 800)

        self.data1 = data1
        self.data2 = data2
        self.name1 = name1 or "控制价"
        self.name2 = name2 or "投标文件"
        self.result = None
        self.aligned1 = self.aligned2 = None

        # 左右树
        self.tree1 = QTreeView()
        self.tree1.setContentsMargins(0,0,0,0)
        self.tree2 = QTreeView()
        self.tree2.setContentsMargins(0,0,0,0)
        for t in (self.tree1, self.tree2):
            t.setHeaderHidden(True)
            t.setDragEnabled(True)
            t.setAcceptDrops(True)
            t.setDropIndicatorShown(True)
            t.setDragDropMode(QTreeView.DragDropMode.InternalMove)
            t.setSelectionMode(QTreeView.SelectionMode.SingleSelection)

        self.tree1.setModel(self.build_model(data1))
        self.tree2.setModel(self.build_model(data2))
        self.tree1.expandAll()
        self.tree2.expandAll()

        # 标题
        # lbl1 = QLabel(f"{self.name1}")
        # lbl2 = QLabel(f"投标文件：{self.name2}")

        # 按钮
        btn_align = QPushButton("对齐完成")
        btn_skip = QPushButton("跳过此文件")
        btn_skip.clicked.connect(self.on_skip)
        btn_align.clicked.connect(self.on_align)


        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_box.addWidget(btn_skip)
        btn_box.addStretch()
        btn_box.addWidget(btn_align)
        btn_box.addStretch()

        # 布局
        left = QVBoxLayout()
        # left.addWidget(lbl1)
        left.addWidget(self.tree1)

        right = QVBoxLayout()
        # right.addWidget(lbl2)
        right.addWidget(self.tree2)

        splitter = QSplitter()
        splitter.addWidget(self.wrap_group("招标清单", left))
        splitter.addWidget(self.wrap_group(f"{name2}投标清单", right))
        splitter.setSizes([725, 725])

        main = QVBoxLayout()
        main.addWidget(splitter)
        main.addLayout(btn_box)
        main.addSpacing(20)

        container = QWidget()
        container.setLayout(main)
        self.setCentralWidget(container)

    def wrap_group(self, title, layout):
        w = QWidget()
        w.setLayout(layout)
        box = QGroupBox(title)
        vbox = QVBoxLayout()
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.addWidget(w)
        box.setLayout(vbox)
        return box

    def build_model(self, data):
        model = QStandardItemModel()
        root = model.invisibleRootItem()

        for item in data.get("单项工程", []):
            name = item.get("名称", "未知单项")
            inode = QStandardItem(name)
            font = QFont()
            font.setPointSize(11)
            font.setBold(True)
            inode.setFont(font)
            inode.setEditable(False)
            root.appendRow(inode)

            for unit in item.get("单位工程", []):
                uname = unit.get("名称", "未知单位")
                unode = QStandardItem(uname)
                unode.setEditable(False)
                inode.appendRow(unode)
        return model

    def model_to_data(self, model, src_data):
        result = {"单项工程": []}
        root = model.invisibleRootItem()
        for i in range(root.rowCount()):
            item_node = root.child(i)
            item_name = item_node.text().split("：", 1)[1]
            item_dict = {"名称": item_name, "单位工程": []}
            for j in range(item_node.rowCount()):
                unit_node = item_node.child(j)
                unit_name = unit_node.text().split("：", 1)[1].strip()
                orig = self.find_original_unit(src_data, item_name, unit_name)
                item_dict["单位工程"].append(orig)
            result["单项工程"].append(item_dict)
        return result

    def find_original_unit(self, src_data, item_name, unit_name):
        for item in src_data.get("单项工程", []):
            if item.get("名称") == item_name:
                for unit in item.get("单位工程", []):
                    if unit.get("名称") == unit_name:
                        return unit
        return {"名称": unit_name, "分部清单": []}

    def on_align(self):
        if QMessageBox.question(self, "确认对齐",
                                "左右结构已完全一致？\n点击“是”后将使用当前顺序",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                QMessageBox.StandardButton.Yes) == QMessageBox.StandardButton.Yes:
            self.aligned1 = self.model_to_data(self.tree1.model(), self.data1)
            self.aligned2 = self.model_to_data(self.tree2.model(), self.data2)
            self.result = "aligned"
            self.close()

    def on_skip(self):
        if QMessageBox.question(self, "跳过文件",
                                "确定要跳过此投标文件的导入吗？",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self.result = "skip"
            self.close()


def safe_zip_compare(data1, data2):
    """尝试直接zip对比，成功返回True，失败返回False"""
    try:
        for item1, item2 in zip(data1["单项工程"], data2["单项工程"]):
            for u1, u2 in zip(item1["单位工程"], item2["单位工程"]):
                if u1["名称"] == u2["名称"]:
                    continue
                else:
                    print(f"不一致项：{item1['名称']} {u1['名称']}")
                    return False
        return True
    except Exception as e:
        print(f"直接对比失败: {e}")
        return False


def force_align_until_success(data1, data2, name1="文件1", name2="文件2"):
    """反复弹窗，直到完全对齐或用户跳过/放弃"""
    current1, current2 = data1, data2

    while True:
        # 1. 尝试直接 zip 对比
        if safe_zip_compare(current1, current2):
            print("结构完全一致，对比成功！")
            return current1, current2, False  # False 表示没有跳过

        # 2. 不一致 → 弹出新版窗口（带跳过按钮）
        print("结构不一致，启动手动对齐工具...")
        app = QApplication.instance() or QApplication(sys.argv)
        win = StructureAligner(current1, current2, name1, name2)
        win.show()
        app.exec()

        # 3. 根据用户选择处理
        if win.result == "aligned":
            current1, current2 = win.aligned1, win.aligned2
            print("已保存调整顺序，准备再次校验...")
            continue

        elif win.result == "skip":
            print("用户选择跳过此文件对比")
            return None, None, True  # True 表示跳过

        else:  # 窗口被直接关闭
            reply = QMessageBox.question(
                None, "未操作",
                "未完成对齐，要放弃整个对比吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                print("用户放弃对比，程序退出")
                sys.exit(0)
            else:
                print("继续调整...")
                continue


