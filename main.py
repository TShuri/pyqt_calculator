import os
import sys

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from logic import Logic


class CalculatorWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt6")
        self.resize(700, 500)
        self.setAcceptDrops(True)
        self.logic = Logic(output_func=self.append_text, ask_gp_callback=self.ask_gp_callback)
        self.files = [] # список для хранения выбранных файлов
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.btn_select_file = QPushButton("Выбрать файлы", self)
        self.btn_select_file.clicked.connect(self.on_select_file)
        layout.addWidget(self.btn_select_file)

        self.btn_run = QPushButton("Рассчитать", self)
        self.btn_run.clicked.connect(self.run)
        layout.addWidget(self.btn_run)

        self.btn_reset = QPushButton("Сброс", self)
        self.btn_reset.clicked.connect(self.on_reset)
        layout.addWidget(self.btn_reset)

        self.text_output = QTextEdit()
        self.text_output.setReadOnly(True)
        self.text_output.setPlaceholderText("Перетащите сюда файлы Excel или выберите их кнопкой Выбрать файлы")
        layout.addWidget(self.text_output)

        self.setLayout(layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            new_files = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
            
            # оставляем только те файлы, которых ещё нет
            unique_new_files = [f for f in new_files if f not in self.files]
            
            # добавляем в общий список
            if unique_new_files:
                self.files.extend(unique_new_files)
                file_names = [os.path.basename(f) for f in unique_new_files]
                self.append_text("\n".join(file_names))
            
            event.acceptProposedAction()
        else:
            event.ignore()

    def on_select_file(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Выберите файлы Excel", "", "Excel Files (*.xlsx)")
        if files:
            # оставляем только те файлы, которых ещё нет
            unique_new_files = [f for f in files if f not in self.files]
            
            if unique_new_files:
                self.files.extend(unique_new_files)
                file_names = [os.path.basename(f) for f in unique_new_files]
                self.append_text("\n".join(file_names))
    
    def on_reset(self):
        self.text_output.clear()
        self.files = []

    def append_text(self, msg):
        self.text_output.append(str(msg))

    def ask_gp_callback(self, msg):
        reply = QMessageBox.question(
            self, "Госпошлина", msg, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        return reply == QMessageBox.StandardButton.Yes

    def run(self):
        if not self.files:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите файлы для расчета.")
            return
        self.logic.run(files=self.files)
        self.append_text("Расчет ...")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CalculatorWindow()
    window.show()
    sys.exit(app.exec())
