import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QPushButton,
    QFileDialog, QVBoxLayout, QHBoxLayout, QWidget, QMessageBox, QListWidget
)
from PyQt5.QtCore import Qt
from converters.word2md_converter import Word2MarkdownConverter
from parsers.document_parser import WordDocumentParser


class Word2MdGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Word to Markdown Converter @vrxiaojie")
        self.setGeometry(100, 100, 800, 400)

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Input file selection
        self.input_label = QLabel("选择输入的 Word 文件：")
        layout.addWidget(self.input_label)

        self.input_line = QLineEdit(self)
        layout.addWidget(self.input_line)

        self.input_button = QPushButton("浏览", self)
        self.input_button.clicked.connect(self.select_input_file)
        layout.addWidget(self.input_button)

        # Section selection and conversion list
        list_layout = QHBoxLayout()

        self.all_sections_list = QListWidget(self)
        self.all_sections_list.setSelectionMode(QListWidget.MultiSelection)
        list_layout.addWidget(self.all_sections_list)

        button_layout = QVBoxLayout()
        self.add_button = QPushButton("添加到转换列表", self)
        self.add_button.clicked.connect(self.add_to_conversion_list)
        button_layout.addWidget(self.add_button)

        self.remove_button = QPushButton("从转换列表移除", self)
        self.remove_button.clicked.connect(self.remove_from_conversion_list)
        button_layout.addWidget(self.remove_button)

        button_layout.addStretch()
        list_layout.addLayout(button_layout)

        self.selected_sections_list = QListWidget(self)
        self.selected_sections_list.setSelectionMode(QListWidget.MultiSelection)
        list_layout.addWidget(self.selected_sections_list)

        layout.addLayout(list_layout)

        # Output file selection
        self.output_label = QLabel("选择输出的 Markdown 文件：")
        layout.addWidget(self.output_label)

        self.output_line = QLineEdit(self)
        layout.addWidget(self.output_line)

        self.output_button = QPushButton("浏览", self)
        self.output_button.clicked.connect(self.select_output_file)
        layout.addWidget(self.output_button)

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
                if section['index']>0:
                    self.all_sections_list.addItem(f"{section['index']}: {section['title']}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"解析 Word 文件时发生错误：{str(e)}")

    def add_to_conversion_list(self):
        """将选中的段落添加到待转换列表"""
        selected_items = self.all_sections_list.selectedItems()
        for item in selected_items:
            if item.text() not in [self.selected_sections_list.item(i).text() for i in range(self.selected_sections_list.count())]:
                self.selected_sections_list.addItem(item.text())

    def remove_from_conversion_list(self):
        """从待转换列表中移除选中的段落"""
        selected_items = self.selected_sections_list.selectedItems()
        for item in selected_items:
            self.selected_sections_list.takeItem(self.selected_sections_list.row(item))

    def select_output_file(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "选择 Markdown 文件", "", "Markdown Files (*.md)")
        if file_name:
            self.output_line.setText(file_name)

    def convert_word_to_md(self):
        input_path = self.input_line.text()
        output_path = self.output_line.text()

        if not input_path or not output_path:
            QMessageBox.warning(self, "警告", "请输入完整的文件路径！")
            return

        try:
            selected_indices = []
            for i in range(self.selected_sections_list.count()):
                text = self.selected_sections_list.item(i).text()
                index = int(text.split(":")[0])  # 提取段落索引
                selected_indices.append(index)

            converter = Word2MarkdownConverter()
            success = converter.convert(input_path, output_path, section_range=selected_indices)

            if success:
                QMessageBox.information(self, "成功", "转换完成！")
            else:
                QMessageBox.warning(self, "失败", "转换失败，请检查输入文件。")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生错误：{str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = Word2MdGUI()
    main_window.show()
    sys.exit(app.exec_())