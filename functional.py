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


def send_email(sender, password, subject, message):
    server = smtplib.SMTP_SSL('smtp.mail.ru:465')

    try:
        server.login(sender, password)
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = subject

        for file in os.listdir("attachments"):
            time.sleep(0.4)
            filename = os.path.basename(file)
            ftype, encoding = mimetypes.guess_type(file)
            file_type, subtype = ftype.split("/")

            if file_type == "text":
                with open(f"attachments/{file}") as f:
                    file = MIMEText(f.read())
            elif file_type == "image":
                with open(f"attachments/{file}", "rb") as f:
                    file = MIMEImage(f.read(), subtype)
            elif file_type == "audio":
                with open(f"attachments/{file}", "rb") as f:
                    file = MIMEAudio(f.read(), subtype)
            elif file_type == "application":
                with open(f"attachments/{file}", "rb") as f:
                    file = MIMEApplication(f.read(), subtype)
            else:
                with open(f"attachments/{file}", "rb") as f:
                    file = MIMEBase(file_type, subtype)
                    file.set_payload(f.read())
                    encoders.encode_base64(file)

            file.add_header('content-disposition', 'attachment', filename=filename)
            msg.attach(file)

        msg.attach(MIMEText(message))
        server.sendmail(msg.as_string())

        return "Сообщение было успешно отправлено!"
    except Exception as _ex:
        return f"{_ex}\nПожалуйста, проверьте свой логин или пароль!"


def main():
    base = openpyxl.open("C:\\Users\\Denis\\PycharmProjects\\pythonProject1\\base\\1.xlsx", read_only=True)
    sheet = base.active

    sender = sheet["A1"].value
    password = sheet["B1"].value

    subject = input("Введите свою тему: ")
    message = input("Введите свое сообщение: ")

    print(send_email(sender=sender, password=password, subject=subject, message=message))


if __name__ == "__main__":
    main()
