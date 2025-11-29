import sys
import os
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog, QFileDialog, QMessageBox
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QPixmap


class DialogWindow(QDialog):
    def __init__(self):
        super(DialogWindow, self).__init__()
        uic.loadUi('src/dialog_window.ui', self)

        # Connect buttons to their functions
        self.pushButton.clicked.connect(self.download_template)
        self.pushButton_2.clicked.connect(self.upload_table)

    def download_template(self):
        # Function to download template
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить шаблон", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            # Here you would implement the actual template download logic
            QMessageBox.information(self, "Успех", f"Шаблон сохранен как: {file_path}")

    def upload_table(self):
        # Function to upload table
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите таблицу", "",
                                                   "Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)")
        if file_path:
            # Here you would implement the actual table upload logic
            QMessageBox.information(self, "Успех", f"Таблица загружена: {file_path}")


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('src/main.ui', self)

        # Set current date
        self.dateEdit.setDate(QDate.currentDate())

        # Connect buttons to their functions
        self.pushButton.clicked.connect(self.open_dialog)
        self.pushButton_2.clicked.connect(self.download_certificates)
        self.pushButton_6.clicked.connect(self.add_custom_template)

        # Load images (you might need to adjust paths)
        self.load_images()

    def load_images(self):
        # Load logo
        if os.path.exists("src/logo.jpg"):
            self.label.setPixmap(QPixmap("src/logo.jpg"))

        # Load other images
        image_paths = [
            ("src/vector.jpg", self.label_2),
            ("src/example1.jpg", self.label_3),
            ("src/example3.jpg", self.label_4),
            ("src/example2.jpg", self.label_5),
            ("src/example4.jpg", self.label_6)
        ]

        for path, label in image_paths:
            if os.path.exists(path):
                label.setPixmap(QPixmap(path))

    def open_dialog(self):
        dialog = DialogWindow()
        dialog.exec_()

    def download_certificates(self):
        # Get selected category
        category = ""
        if self.radioButton.isChecked():
            category = "Наука"
        elif self.radioButton_2.isChecked():
            category = "Искусство"
        elif self.radioButton_3.isChecked():
            category = "Спорт"

        # Get date
        selected_date = self.dateEdit.date().toString("dd.MM.yyyy")

        # Get event name
        event_name = self.textEdit.toPlainText()

        # Validate inputs
        if not category:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите категорию")
            return

        if not event_name.strip():
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите название мероприятия")
            return

        # Here you would implement the actual certificate generation and download logic
        QMessageBox.information(self, "Успех",
                                f"Сертификаты для мероприятия '{event_name}' ({category}) от {selected_date} готовы к скачиванию!")

        # Open file dialog to save certificates
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить сертификаты", f"сертификаты_{event_name}",
                                                   "PDF Files (*.pdf);;All Files (*)")
        if file_path:
            QMessageBox.information(self, "Успех", f"Сертификаты сохранены как: {file_path}")

    def add_custom_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите шаблон", "",
                                                   "Hypertext Markup Language (*.htm *.html);; All files (*.*)")
        if file_path:
            # Here you would implement the custom template addition logic
            QMessageBox.information(self, "Успех", f"Пользовательский шаблон добавлен: {file_path}")


def main():
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
