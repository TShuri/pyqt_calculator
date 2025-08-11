# QT_QPA_PLATFORM=wayland python3 app_pyqt6.py
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QFileDialog, QMessageBox
)
from PyQt6.QtCore import Qt
import sys
from logic import Logic

class CalculatorWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('PyQt6')
        self.resize(700, 500)
        self.logic = Logic(output_func=self.append_text, ask_gp_callback=self.ask_gp_callback)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.btn_select_file = QPushButton('Выбрать файлы', self)
        self.btn_select_file.clicked.connect(self.on_select_file)
        layout.addWidget(self.btn_select_file)
        
        self.btn_reset = QPushButton('Сброс', self)
        self.btn_reset.clicked.connect(self.on_reset)
        layout.addWidget(self.btn_reset)

        self.text_output = QTextEdit()
        self.text_output.setReadOnly(True)
        layout.addWidget(self.text_output)

        self.setLayout(layout)
            
    def on_select_file(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Выберите файлы Excel", "", "Excel Files (*.xlsx)")
        if files:
            self.on_reset()  
            QMessageBox.information(self, "Файлы выбраны", f"Выбрано файлов: {len(files)}")
            self.logic.run(files=files)
        else:
            QMessageBox.information(self, "Предупреждение", "Файлы не выбраны.")
            self.append_text("Файлы не выбраны.")
    
    def on_reset(self):
        self.text_output.clear()
        
    def append_text(self, msg):
        self.text_output.append(str(msg))
        
    def ask_gp_callback(self, msg):
        reply = QMessageBox.question(
            self,
            "Госпошлина",
            msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        return reply == QMessageBox.StandardButton.Yes
                
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CalculatorWindow()
    window.show()
    sys.exit(app.exec())
