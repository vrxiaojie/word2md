def get_stylesheet():
    return """
        QMainWindow {
            background-color: #f5f5f5;
        }
        QLabel {
            font-size: 14px;
            color: #333;
        }
        QLineEdit {
            border: 1px solid #ccc;
            border-radius: 5px;
            padding: 5px;
            font-size: 14px;
        }
        QPushButton {
            background-color: #0078d7;
            color: white;
            border: none;
            border-radius: 5px;
            padding: 8px 15px;
            font-size: 14px;
        }
        QPushButton:hover {
            background-color: #005a9e;
        }
        QListWidget {
            border: 1px solid #ccc;
            border-radius: 5px;
            font-size: 14px;
        }
        QComboBox {
            border: 1px solid #ccc;
            border-radius: 5px;
            padding: 5px;
            font-size: 14px;
        }
    """