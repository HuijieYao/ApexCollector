import base64
import os
import winreg

import globalValue as Pb
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QListWidget, QLabel, QLineEdit, QMainWindow, \
    QPushButton, QComboBox, QFileDialog, QRadioButton

from images.setting_png import img as setting_png


class ChildWindow(QMainWindow):
    signal = QtCore.pyqtSignal(bool)

    def __init__(self):
        super().__init__()
        self.list_menu = QListWidget()
        self.list_menu.setFixedWidth(150)
        self.list_menu.addItems(["常规", "高级"])

        self.dir_label = QLabel("保存路径")
        self.apart_label = QLabel("楼栋编号")
        self.auto_label = QLabel("是否开启自动打包")

        self.dir_edit = QLineEdit()
        self.apart_edit = QLineEdit()

        self.floor_edit = QComboBox(self)
        self.floor_edit.addItems(['选择楼层', '1楼', '2楼', '3楼', '4楼', '5楼', '6楼'])

        self.v_spacer = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.h_spacer = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)

        self.dir_button = QPushButton("选择路径")
        self.certain_button = QPushButton("确定")
        self.cancel_button = QPushButton("取消")
        self.apply_button = QPushButton("应用")

        self.radio_button_open = QRadioButton("开启", self)
        self.radio_button_close = QRadioButton("关闭", self)

        self.frame_1 = QtWidgets.QFrame()
        self.frame_2 = QtWidgets.QFrame()

        self.init_gui()

    def init_gui(self):
        """
        初始化MainWindow
        :return: null
        """
        self.setFixedSize(480, 270)
        self.setWindowTitle("设置")

        tmp = open("setting.png", "wb")
        tmp.write(base64.b64decode(setting_png))
        tmp.close()
        self.setWindowIcon(QIcon(r"setting.png"))
        os.remove("setting.png")

        self.init_layout()
        self.add_affairs()
        self.setup()

    def init_layout(self):
        """
        设置布局
        :return: null
        """
        # 创建全局布局
        global_layout = QHBoxLayout()

        # 创建局部布局
        left_v_layout = QVBoxLayout()
        right_v_layout_1 = QVBoxLayout()
        right_v_layout_2 = QVBoxLayout()
        right_v_layout = QVBoxLayout()

        left_v_layout.addWidget(self.list_menu)

        right_dir_h_layout = QHBoxLayout()
        right_apart_h_layout = QHBoxLayout()
        right_auto_h_layout = QHBoxLayout()
        right_option_h_layout = QHBoxLayout()

        right_dir_h_layout.addWidget(self.dir_label)
        right_dir_h_layout.addWidget(self.dir_edit)
        right_dir_h_layout.addWidget(self.dir_button)
        right_v_layout_1.addLayout(right_dir_h_layout)

        right_apart_h_layout.addWidget(self.apart_label)
        right_apart_h_layout.addWidget(self.apart_edit)
        right_apart_h_layout.addWidget(self.floor_edit)
        right_v_layout_1.addLayout(right_apart_h_layout)

        right_v_layout_1.addItem(self.v_spacer)

        right_option_h_layout.addItem(self.h_spacer)
        right_option_h_layout.addWidget(self.certain_button)
        right_option_h_layout.addWidget(self.cancel_button)
        right_option_h_layout.addWidget(self.apply_button)

        right_auto_h_layout.addWidget(self.auto_label)
        right_auto_h_layout.addWidget(self.radio_button_open)
        right_auto_h_layout.addWidget(self.radio_button_close)
        right_v_layout_2.addLayout(right_auto_h_layout)
        right_v_layout_2.addItem(self.v_spacer)

        self.frame_1.setLayout(right_v_layout_1)
        self.frame_2.setLayout(right_v_layout_2)

        right_v_layout.addWidget(self.frame_1)
        right_v_layout.addWidget(self.frame_2)
        right_v_layout.addLayout(right_option_h_layout)

        # 局部布局添加到全局布局
        global_layout.addLayout(left_v_layout)
        global_layout.addLayout(right_v_layout)
        self.frame_1.show()
        self.frame_2.hide()
        global_layout.setStretch(0, 1)
        global_layout.setStretch(1, 2)

        widget = QWidget()
        widget.setLayout(global_layout)
        self.setCentralWidget(widget)

    def slot(self):
        self.signal.emit(Pb.is_changed)

    def setup(self):
        self.dir_edit.setText(Pb.dir_path)
        self.apart_edit.setText(Pb.apart)
        self.floor_edit.setCurrentIndex(Pb.floor_index)
        if Pb.auto_pack == "true":
            self.radio_button_open.setChecked(True)
        else:
            self.radio_button_close.setChecked(True)

    def clicked(self, item):
        if item.text() == "常规":
            self.frame_1.show()
            self.frame_2.hide()
        elif item.text() == "高级":
            self.frame_1.hide()
            self.frame_2.show()

    def click_choice_dir(self):
        """
        获取保存结果的路径
        :return: null
        """
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        dir_path = QFileDialog.getExistingDirectory(self, "请选择文件夹路径", winreg.QueryValueEx(key, "Desktop")[0])
        self.dir_edit.setText(dir_path)

    def click_certain(self):
        self.click_apply()
        self.close()

    def click_apply(self):
        if self.frame_1.isVisible():
            Pb.set_dir_path(self.dir_edit.text())
            Pb.set_apart(self.apart_edit.text())
            Pb.set_floor_index(self.floor_edit.currentIndex())
        else:
            Pb.set_auto_pack(self.radio_button_open.isChecked())

    def closeEvent(self, event):
        self.slot()
        self.close()

    def add_affairs(self):
        self.list_menu.itemClicked.connect(self.clicked)
        self.dir_button.clicked.connect(lambda: self.click_choice_dir())
        self.certain_button.clicked.connect(lambda: self.click_certain())
        self.cancel_button.clicked.connect(self.closeEvent)
        self.apply_button.clicked.connect(lambda: self.click_apply())
