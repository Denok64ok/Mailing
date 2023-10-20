import smtplib
import os
import time
import mimetypes

from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase

import openpyxl
from PyQt5.QtWidgets import QMessageBox, QApplication


class Mail:
    def __init__(self, from_directories, to_directories, files_directories, subject, message):
        self.from_directories = from_directories
        self.to_directories = to_directories
        self.files_directories = files_directories
        self.subject = subject
        self.message = message

    def send_email(self):
        server = smtplib.SMTP_SSL('smtp.mail.ru:465')

        try:
            from_base = openpyxl.open(self.from_directories, read_only=True)
            from_sheet = from_base.active

            sender = from_sheet["A1"].value
            password = from_sheet["B1"].value

            server.login(sender, password)
            msg = MIMEMultipart()
            msg["Subject"] = self.subject

            for file in os.listdir(self.files_directories):
                time.sleep(0.4)
                filename = os.path.basename(file)
                ftype, encoding = mimetypes.guess_type(file)
                file_type, subtype = ftype.split("/")

                if file_type == "text":
                    with open(self.files_directories + "/" + file) as f:
                        file = MIMEText(f.read())
                elif file_type == "image":
                    with open(self.files_directories + "/" + file, "rb") as f:
                        file = MIMEImage(f.read(), subtype)
                elif file_type == "audio":
                    with open(self.files_directories + "/" + file, "rb") as f:
                        file = MIMEAudio(f.read(), subtype)
                elif file_type == "application":
                    with open(self.files_directories + "/" + file, "rb") as f:
                        file = MIMEApplication(f.read(), subtype)
                else:
                    with open(self.files_directories + "/" + file, "rb") as f:
                        file = MIMEBase(file_type, subtype)
                        file.set_payload(f.read())
                        encoders.encode_base64(file)

                file.add_header('content-disposition', 'attachment', filename=filename)
                msg.attach(file)

            msg.attach(MIMEText(self.message))

            to_base = openpyxl.open(self.to_directories, read_only=True)
            to_sheet = to_base.active

            i = 1

            error = QMessageBox()
            error.setWindowTitle("Внимание")
            error.setIcon(QMessageBox.Warning)
            error.setText("Производится почтовая рассылка.\nПожалуйста подождите.")
            error.show()
            QApplication.processEvents()

            while to_sheet["A" + str(i)].value != None:
                time.sleep(0.4)
                server.sendmail(from_addr=sender, to_addrs=to_sheet["A" + str(i)].value, msg=msg.as_string())
                i += 1

            # return "Сообщение было успешно отправлено!"
            error.close()
        except Exception as _ex:
            # return f"{_ex}\nПожалуйста, проверьте свой логин или пароль!"
            error = QMessageBox()
            error.setWindowTitle("Произошла ощибка")
            error.setIcon(QMessageBox.Warning)
            error.setText("Пожалуйста, проверьте свой логин или пароль,  а также пути до необходимых файлов!")
            error.exec_()
