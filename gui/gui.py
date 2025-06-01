import os
import sys

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QPushButton,
    QFileDialog, QVBoxLayout, QHBoxLayout, QWidget, QMessageBox, QListWidget, QComboBox
)
from converters.word2md_converter import Word2MarkdownConverter
from parsers.document_parser import WordDocumentParser
from style import get_stylesheet  # 导入样式


class Word2MdGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Word to Markdown Converter @vrxiaojie")
        self.setGeometry(100, 100, 900, 500)

        self.setStyleSheet(get_stylesheet())  # 应用样式表
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Input file selection
        input_layout = QHBoxLayout()
        self.input_label = QLabel("选择输入的 Word 文件：")
        input_layout.addWidget(self.input_label)

        self.input_line = QLineEdit(self)
        input_layout.addWidget(self.input_line)

        self.input_button = QPushButton("浏览", self)
        self.input_button.clicked.connect(self.select_input_file)
        input_layout.addWidget(self.input_button)

        layout.addLayout(input_layout)

        # Section selection and conversion list
        list_layout = QHBoxLayout()

        self.all_sections_list = QListWidget(self)
        self.all_sections_list.setSelectionMode(QListWidget.MultiSelection)
        list_layout.addWidget(self.all_sections_list)

        button_layout = QVBoxLayout()
        button_layout.addStretch()
        self.add_button = QPushButton("添加到转换列表", self)
        self.add_button.clicked.connect(self.add_to_conversion_list)
        button_layout.addWidget(self.add_button)

        self.add_all_button = QPushButton("全部加入", self)  # 全部加入按钮
        self.add_all_button.clicked.connect(self.add_all_to_conversion_list)
        button_layout.addWidget(self.add_all_button)

        self.remove_button = QPushButton("从转换列表移除", self)
        self.remove_button.clicked.connect(self.remove_from_conversion_list)
        button_layout.addWidget(self.remove_button)

        self.remove_all_button = QPushButton("全部移除", self)  # 全部移除按钮
        self.remove_all_button.clicked.connect(self.remove_all_from_conversion_list)
        button_layout.addWidget(self.remove_all_button)

        button_layout.addStretch()

        # Code block language selection
        lang_layout = QHBoxLayout()
        self.lang_label = QLabel("代码块语言：")
        lang_layout.addWidget(self.lang_label)

        self.lang_combo = QComboBox(self)
        self.lang_combo.addItems(["", "Python", "C", "Java", "JavaScript", "HTML", "CSS"])
        lang_layout.addWidget(self.lang_combo)
        button_layout.addLayout(lang_layout)

        list_layout.addLayout(button_layout)

        self.selected_sections_list = QListWidget(self)
        self.selected_sections_list.setSelectionMode(QListWidget.MultiSelection)
        list_layout.addWidget(self.selected_sections_list)

        layout.addLayout(list_layout)

        # Output file selection
        output_layout = QHBoxLayout()
        self.output_label = QLabel("选择输出的 Markdown 文件：")
        output_layout.addWidget(self.output_label)

        self.output_line = QLineEdit(self)
        output_layout.addWidget(self.output_line)

        self.output_button = QPushButton("浏览", self)
        self.output_button.clicked.connect(self.select_output_file)
        output_layout.addWidget(self.output_button)

        layout.addLayout(output_layout)

        # Convert button
        self.convert_button = QPushButton("转换", self)
        self.convert_button.clicked.connect(self.convert_word_to_md)
        layout.addWidget(self.convert_button)

        # Setting the central widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_input_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "选择 Word 文件", "", "Word Files (*.docx)")
        if file_name:
            self.input_line.setText(file_name)
            self.load_sections(file_name)

    def load_sections(self, file_name):
        """加载并展示 Word 文件中的段落"""
        try:
            parser = WordDocumentParser(file_name)
            self.all_sections_list.clear()
            for section in parser.get_sections():
                if section['index'] > 0:
                    self.all_sections_list.addItem(f"{section['index']}: {section['title']}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"解析 Word 文件时发生错误：{str(e)}")

    def add_to_conversion_list(self):
        """将选中的段落添加到待转换列表"""
        selected_items = self.all_sections_list.selectedItems()
        for item in selected_items:
            if item.text() not in [self.selected_sections_list.item(i).text() for i in range(self.selected_sections_list.count())]:
                self.selected_sections_list.addItem(item.text())

    def add_all_to_conversion_list(self):
        """将所有段落添加到待转换列表"""
        for i in range(self.all_sections_list.count()):
            item_text = self.all_sections_list.item(i).text()
            if item_text not in [self.selected_sections_list.item(j).text() for j in range(self.selected_sections_list.count())]:
                self.selected_sections_list.addItem(item_text)

    def remove_from_conversion_list(self):
        """从待转换列表中移除选中的段落"""
        selected_items = self.selected_sections_list.selectedItems()
        for item in selected_items:
            self.selected_sections_list.takeItem(self.selected_sections_list.row(item))

    def remove_all_from_conversion_list(self):
        """清空待转换列表"""
        self.selected_sections_list.clear()

    def select_output_file(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "选择 Markdown 文件", "", "Markdown Files (*.md)")
        if file_name:
            self.output_line.setText(file_name)

    def convert_word_to_md(self):
        input_path = self.input_line.text()
        output_path = self.output_line.text()
        lang = self.lang_combo.currentText()

        if not input_path or not output_path:
            QMessageBox.warning(self, "警告", "请输入完整的文件路径！")
            return

        try:
            selected_indices = []
            for i in range(self.selected_sections_list.count()):
                text = self.selected_sections_list.item(i).text()
                index = int(text.split(":")[0])  # 提取段落索引
                selected_indices.append(index)

            if not selected_indices:
                QMessageBox.warning(self, "警告", "没有选择任何章节")
                return

            converter = Word2MarkdownConverter(code_language=lang)
            success = converter.convert(input_path, output_path, section_range=selected_indices)

            if success:
                self.show_success_message(output_path)
            else:
                QMessageBox.warning(self, "失败", "转换失败，请检查输入文件。")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生错误：{str(e)}")

    def show_success_message(self, output_path):
        """显示成功提示框，并添加两个按钮"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("转换成功")
        msg_box.setText("✅转换完成！")
        msg_box.setIcon(QMessageBox.Information)

        open_file_btn = msg_box.addButton("打开 Markdown 文件", QMessageBox.ActionRole)
        open_folder_btn = msg_box.addButton("打开输出文件夹", QMessageBox.ActionRole)
        msg_box.addButton("确定", QMessageBox.RejectRole)

        msg_box.exec_()

        if msg_box.clickedButton() == open_file_btn:
            os.startfile(output_path)  # 打开 Markdown 文件
        elif msg_box.clickedButton() == open_folder_btn:
            folder_path = os.path.dirname(output_path)
            os.startfile(folder_path)  # 打开输出文件夹


def get_resource_path(relative_path):
    """获取资源文件的绝对路径（支持 PyInstaller 打包后的路径）"""
    # 如果程序是打包后的状态，会有 _MEIPASS 属性
    if hasattr(sys, '_MEIPASS'):
        # 资源文件位于 _internal 文件夹中
        return os.path.join(sys._MEIPASS, relative_path)
    # 开发环境直接使用相对路径
    return os.path.join(os.path.abspath("."), relative_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = Word2MdGUI()
    main_window.setWindowIcon(QIcon(get_resource_path("logo.ico")))
    main_window.show()
    sys.exit(app.exec_())