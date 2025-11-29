import sys
import os
import shutil
from PyQt5 import QtWidgets, uic, QtCore
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog, QFileDialog, QMessageBox
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QPixmap, QImage
import pandas as pd

table_path = ''


class SaveDialogWindow(QDialog):
    def __init__(self, parent=None, certificates_data=None):
        super(SaveDialogWindow, self).__init__(parent)
        uic.loadUi('src/save_dialog.ui', self)

        self.certificates_data = certificates_data or []
        self.main_window = parent

        # Connect buttons
        self.pushButton.clicked.connect(self.save_as_pdf)
        self.pushButton_2.clicked.connect(self.save_and_send)

        # Generate and display preview
        self.update_preview()

    def update_preview(self):
        """Generate and display PDF preview"""
        if not self.certificates_data:
            self.show_placeholder("Нет данных сертификатов")
            return

        try:
            # Generate actual PDF for the first participant
            cert_data = self.certificates_data[0]
            pdf_path = self.generate_preview_pdf(cert_data)

            if pdf_path and os.path.exists(pdf_path):
                self.display_pdf_preview_simple(pdf_path)
            else:
                self.show_placeholder("Не удалось сгенерировать предпросмотр")

        except Exception as e:
            print(f"Preview error: {e}")
            self.show_placeholder(f"Ошибка предпросмотра: {str(e)}")

    def generate_preview_pdf(self, cert_data):
        """Generate actual PDF for preview"""
        try:
            from app import render_pdf

            event_info = {
                "name": cert_data["event_name"],
                "date": cert_data["event_date"],
                "track": cert_data["event_category"]
            }

            participant_info = {
                "name": cert_data["participant_name"],
                "surname": cert_data["participant_surname"],
                "email": cert_data["participant_email"],
                "language": cert_data["language"],
                "role": cert_data["participant_role"],
                "place": cert_data["participant_place"]
            }

            # Generate actual PDF
            pdf_path = render_pdf(event_info, participant_info)
            return pdf_path

        except Exception as e:
            print(f"PDF generation error: {e}")
            return None

    def display_pdf_preview_simple(self, pdf_path):
        """Simple PDF preview using PyMuPDF without complex rendering"""
        try:
            import fitz  # PyMuPDF

            # Open PDF document
            doc = fitz.open(pdf_path)

            # Get first page
            page = doc[0]

            # Create matrix for reasonable quality
            mat = fitz.Matrix(1.5, 1.5)  # 1.5x zoom

            # Render page to pixmap
            pix = page.get_pixmap(matrix=mat)

            # Convert to image data (PNG format)
            img_data = pix.tobytes("png")

            # Convert to QPixmap
            pixmap = QPixmap()
            pixmap.loadFromData(img_data)

            # Close document
            doc.close()

            if not pixmap.isNull():
                # Scale to fit graphics view
                scaled_pixmap = pixmap.scaled(
                    self.graphicsView.width() - 20,
                    self.graphicsView.height() - 20,
                    QtCore.Qt.KeepAspectRatio,
                    QtCore.Qt.SmoothTransformation
                )

                # Display in graphics view
                scene = QtWidgets.QGraphicsScene()
                scene.addPixmap(scaled_pixmap)
                self.graphicsView.setScene(scene)
                self.graphicsView.fitInView(scene.itemsBoundingRect(), QtCore.Qt.KeepAspectRatio)
            else:
                self.show_placeholder("Не удалось загрузить изображение предпросмотра")

        except ImportError:
            self.show_placeholder("Установите PyMuPDF: pip install PyMuPDF")
        except Exception as e:
            print(f"Display error: {e}")
            self.show_placeholder("Ошибка отображения PDF")

    def show_placeholder(self, message):
        """Show placeholder message"""
        scene = QtWidgets.QGraphicsScene()
        text_item = QtWidgets.QGraphicsTextItem(message)
        scene.addItem(text_item)
        self.graphicsView.setScene(scene)

    def save_as_pdf(self):
        """Save all certificates as PDF files in selected directory"""
        if not self.certificates_data:
            QMessageBox.warning(self, "Ошибка", "Нет данных сертификатов для сохранения")
            return

        try:
            directory = QFileDialog.getExistingDirectory(
                self,
                "Выберите папку для сохранения сертификатов",
                "",
                QFileDialog.ShowDirsOnly
            )

            if directory:
                saved_files = self.save_certificates_to_directory(directory)
                QMessageBox.information(self, "Успех",
                                        f"Создано {len(saved_files)} сертификатов в папке:\n{directory}")
                self.accept()

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении: {str(e)}")

    def save_certificates_to_directory(self, directory):
        """Save all certificates to the specified directory"""
        saved_files = []

        for cert_data in self.certificates_data:
            try:
                # Generate PDF for each certificate
                pdf_path = self.generate_certificate_pdf(cert_data)

                # Move PDF to selected directory
                if pdf_path and os.path.exists(pdf_path):
                    filename = os.path.basename(pdf_path)
                    new_path = os.path.join(directory, filename)

                    # If file already exists in destination, remove it first
                    if os.path.exists(new_path):
                        os.remove(new_path)

                    shutil.move(pdf_path, new_path)
                    saved_files.append(new_path)

            except Exception as e:
                print(f"Error saving certificate for {cert_data['participant_name']}: {str(e)}")
                continue

        return saved_files

    def generate_certificate_pdf(self, cert_data):
        """Generate PDF for a certificate"""
        from app import render_pdf

        event_info = {
            "name": cert_data["event_name"],
            "date": cert_data["event_date"],
            "track": cert_data["event_category"]
        }

        participant_info = {
            "name": cert_data["participant_name"],
            "surname": cert_data["participant_surname"],
            "email": cert_data["participant_email"],
            "language": cert_data["language"],
            "role": cert_data["participant_role"],
            "place": cert_data["participant_place"]
        }

        return render_pdf(event_info, participant_info)

    def save_and_send(self):
        """Save certificates and start email distribution"""
        if not self.certificates_data:
            QMessageBox.warning(self, "Ошибка", "Нет данных сертификатов для сохранения и отправки")
            return

        try:
            directory = QFileDialog.getExistingDirectory(
                self,
                "Выберите папку для сохранения сертификатов перед отправкой",
                "",
                QFileDialog.ShowDirsOnly
            )

            if directory:
                # Save certificates first
                saved_files = self.save_certificates_to_directory(directory)

                # Start email distribution
                success_count = self.start_email_distribution(saved_files)

                QMessageBox.information(self, "Успех",
                                        f"Создано {len(saved_files)} сертификатов\n"
                                        f"Отправлено {success_count} писем")
                self.accept()

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении и отправке: {str(e)}")

    def start_email_distribution(self, certificate_files):
        """Start sending certificates via email"""
        success_count = 0
        try:
            import mail_handler

            for cert_file in certificate_files:
                try:
                    # Find matching certificate data
                    for cert_data in self.certificates_data:
                        # Match by participant name and event
                        if (cert_data["participant_name"] in cert_file and
                                cert_data["participant_surname"] in cert_file and
                                cert_data["event_name"] in cert_file):
                            mail_handler.send_gmail(
                                cert_data["participant_email"],
                                cert_data["participant_name"],
                                cert_data["language"].lower(),
                                cert_data["event_name"],
                                cert_file
                            )
                            success_count += 1
                            break
                except Exception as e:
                    print(f"Error sending email for {cert_file}: {str(e)}")
                    continue

        except Exception as e:
            print(f"Email distribution error: {e}")

        return success_count

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
        self.generated_certificates = []  # Store generated certificates data
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

            if not layout_added:
                self.label_table_name.setGeometry(80, 630, 400, 30)  # Adjust as needed

    def set_uploaded_table_name(self, table_name):
        if hasattr(self, 'label_table_name'):
            self.label_table_name.setText(f"Загруженная таблица: {table_name}")
            self.label_table_name.setStyleSheet("QLabel { color: green; font-weight: bold; }")
        else:
            self.pushButton.setText(f"Диалог (Таблица: {table_name})")

    def load_images(self):
        if os.path.exists("src/logo.png"):
            self.label.setPixmap(QPixmap("src/logo.png"))

        image_paths = [
            ("src/vector.png", self.label_2),
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
        category = ""
        if self.radioButton.isChecked():
            category = "Наука"
        elif self.radioButton_2.isChecked():
            category = "Искусство"
        elif self.radioButton_3.isChecked():
            category = "Спорт"

        selected_date = self.dateEdit.date().toString("dd.MM.yyyy")

        event_name = self.textEdit.toPlainText()

        if not category:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите категорию")
            return

        if not event_name.strip():
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите название мероприятия")
            return

        if not self.uploaded_table_path or not os.path.exists(self.uploaded_table_path):
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, сначала загрузите таблицу участников")
            return

        try:
            certificates_data = self.generate_certificates(event_name, category, selected_date,
                                                           self.uploaded_table_path)

            save_dialog = SaveDialogWindow(parent=self, certificates_data=certificates_data)
            save_dialog.exec_()

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при генерации сертификатов: {str(e)}")

    def generate_certificates(self, event_name, category, date, table_path):
        """Generate certificates and return certificate data"""
        try:
            from app import get_participants

            event_info = {
                "name": event_name,
                "date": date,
                "track": category
            }

            participants = get_participants(table_path)
            certificates_data = []

            for participant in participants:
                # Store certificate data for preview and later use
                certificate_data = {
                    "participant_name": participant["name"],
                    "participant_surname": participant["surname"],
                    "participant_email": participant["email"],
                    "event_name": event_name,
                    "event_date": date,
                    "event_category": category,
                    "participant_role": participant["role"],
                    "participant_place": participant["place"],
                    "language": participant["language"]
                }
                certificates_data.append(certificate_data)

            return certificates_data

        except Exception as e:
            raise Exception(f"Certificate data preparation failed: {str(e)}")

    def add_custom_template(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите HTML шаблон",
            "",
            "HTML Files (*.html);;All Files (*)"
        )
        if file_path:
            try:
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