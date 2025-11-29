import sys
import os
import shutil
from PyQt5 import QtWidgets, uic, QtCore
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog, QFileDialog, QMessageBox
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QPixmap
import pandas as pd

table_path = ''

class DialogWindow(QDialog):
    def __init__(self, parent=None):
        super(DialogWindow, self).__init__(parent)
        uic.loadUi('src/dialog_window.ui', self)

        # Store parent reference
        self.main_window = parent

        # Connect buttons to their functions
        self.pushButton.clicked.connect(self.download_template)
        self.pushButton_2.clicked.connect(self.upload_table)

    def upload_table(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Выберите таблицу участников",
                "",
                "Excel Files (*.xlsx);;All Files (*)"
            )

            if file_path:
                if self.validate_excel_file(file_path):
                    if self.main_window:
                        self.main_window.uploaded_table_path = file_path
                        self.main_window.set_uploaded_table_name(os.path.basename(file_path))
                    QMessageBox.information(self, "Успех", f"Таблица загружена: {file_path}")
                    self.accept()
                else:
                    QMessageBox.warning(self, "Ошибка",
                                        "Неверный формат таблицы. Убедитесь, что файл содержит все необходимые колонки.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке таблицы: {str(e)}")

    def validate_excel_file(self, file_path):
        try:
            if not os.path.exists(file_path):
                return False
            try:
                df = pd.read_excel(file_path)
                required_columns = ['Имя', 'Фамилия', 'Email', 'Язык', 'Роль', 'Место']
                missing_columns = [col for col in required_columns if col not in df.columns]

                if missing_columns:
                    print(f"Missing columns: {missing_columns}")
                    return False

                return True

            except Exception as e:
                print(f"Error reading Excel file: {e}")
                return False

        except Exception as e:
            print(f"Validation error: {e}")
            return False

    def download_template(self):
        try:
            source_template = "example_data.xlsx"

            if not os.path.exists(source_template):
                QMessageBox.warning(self, "Ошибка", "Шаблон не найден в системе.")
                return

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить шаблон",
                "шаблон_участников.xlsx",
                "Excel Files (*.xlsx);;All Files (*)"
            )

            if file_path:
                shutil.copy2(source_template, file_path)
                QMessageBox.information(self, "Успех", f"Шаблон сохранен как: {file_path}")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить шаблон: {str(e)}")


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('src/main.ui', self)

        self.dateEdit.setDate(QDate.currentDate())
        self.dateEdit.setDisplayFormat("dd.MM.yyyy")
        self.dateEdit.setReadOnly(False)
        self.dateEdit.setCalendarPopup(True)

        # Simple stylesheet that actually works
        self.dateEdit.setStyleSheet("""
            QDateEdit {
                color: black !important;
                background-color: white !important;
                border: 1px solid #cccccc;
                padding: 5px;
            }
            QDateEdit:focus {
                border: 2px solid #0078D7;
            }
        """)


        self.pushButton.clicked.connect(self.open_dialog)
        self.pushButton_2.clicked.connect(self.download_certificates)
        self.pushButton_6.clicked.connect(self.add_custom_template)

        self.uploaded_table_path = None
        self.init_table_label()
        self.load_images()

    def init_table_label(self):
        if not hasattr(self, 'label_table_name'):
            self.label_table_name = QtWidgets.QLabel(self)
            self.label_table_name.setObjectName("label_table_name")
            self.label_table_name.setText("Таблица не загружена")
            self.label_table_name.setStyleSheet("QLabel { color: gray; font-style: italic; }")
            self.label_table_name.setAlignment(QtCore.Qt.AlignCenter)

            layout_added = False

            layout_names = ['verticalLayout', 'gridLayout', 'horizontalLayout', 'formLayout']
            for layout_name in layout_names:
                if hasattr(self, layout_name):
                    layout = getattr(self, layout_name)
                    layout.insertWidget(1, self.label_table_name)
                    layout_added = True
                    break
            if not layout_added:
                self.label_table_name.setGeometry(50, 100, 400, 30)  # Adjust as needed

    def set_uploaded_table_name(self, table_name):
        if hasattr(self, 'label_table_name'):
            self.label_table_name.setText(f"Загруженная таблица: {table_name}")
            self.label_table_name.setStyleSheet("QLabel { color: green; font-weight: bold; }")
        else:
            self.pushButton.setText(f"Диалог (Таблица: {table_name})")

    def load_images(self):
        if os.path.exists("src/logo.jpg"):
            self.label.setPixmap(QPixmap("src/logo.jpg"))

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
        try:
            dialog = DialogWindow(parent=self)
            dialog.exec_()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть диалоговое окно: {str(e)}")
            print(f"Dialog error: {e}")

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

        # Check if table is uploaded
        if not self.uploaded_table_path or not os.path.exists(self.uploaded_table_path):
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, сначала загрузите таблицу участников")
            return

        try:
            # Generate certificates using the uploaded table
            self.generate_certificates(event_name, category, selected_date, self.uploaded_table_path)

            QMessageBox.information(self, "Успех",
                                    f"Сертификаты для мероприятия '{event_name}' ({category}) от {selected_date} готовы к скачиванию!")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при генерации сертификатов: {str(e)}")

    def generate_certificates(self, event_name, category, date, table_path):
        """Generate certificates using the app.py logic"""
        try:
            # Import your certificate generation functions
            from app import get_participants, render_pdf
            import mail_handler

            # Set up event info
            event_info = {
                "name": event_name,
                "date": date,
                "track": category
            }

            # Get participants from uploaded table
            participants = get_participants(table_path)

            # Generate PDFs for each participant
            for participant in participants:
                pdf_path = render_pdf(event_info, participant)
                print(f"Generated PDF: {pdf_path}")

                # Optional: Send emails (you might want to make this configurable)
                # mail_handler.send_gmail(
                #     participant["email"],
                #     participant["name"],
                #     participant["language"].lower(),
                #     event_info["name"],
                #     pdf_path
                # )

        except Exception as e:
            raise Exception(f"Certificate generation failed: {str(e)}")

    def add_custom_template(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите HTML шаблон",
            "",
            "HTML Files (*.html);;All Files (*)"
        )
        if file_path:
            try:
                # Copy template to templates directory
                template_name = os.path.basename(file_path)
                destination = os.path.join("templates", template_name)
                shutil.copy2(file_path, destination)
                QMessageBox.information(self, "Успех", f"Пользовательский шаблон добавлен: {template_name}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось добавить шаблон: {str(e)}")


def main():
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
