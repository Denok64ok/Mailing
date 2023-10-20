import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QDir
from PyQt5.QtWidgets import QFileDialog

import icon_rc


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 500)
        MainWindow.setMinimumSize(QtCore.QSize(800, 500))
        MainWindow.setMaximumSize(QtCore.QSize(800, 500))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/1681536759_papik-pro-p-pochta-logotip-vektor-2.png"),
                       QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setAnimated(True)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.topic = QtWidgets.QLabel(self.centralwidget)
        self.topic.setGeometry(QtCore.QRect(30, 80, 31, 21))
        self.topic.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        self.topic.setAcceptDrops(False)
        self.topic.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.topic.setObjectName("topic")

        self.message = QtWidgets.QLabel(self.centralwidget)
        self.message.setGeometry(QtCore.QRect(20, 160, 61, 21))
        self.message.setObjectName("message")

        self.message_field = QtWidgets.QTextEdit(self.centralwidget)
        self.message_field.setGeometry(QtCore.QRect(20, 190, 761, 251))
        self.message_field.setObjectName("message_field")

        self.newsletter = QtWidgets.QPushButton(self.centralwidget)
        self.newsletter.setGeometry(QtCore.QRect(360, 460, 75, 23))
        self.newsletter.setObjectName("newsletter")
        self.newsletter.clicked.connect(lambda: self.start_newsletter())

        self.subject_field = QtWidgets.QLineEdit(self.centralwidget)
        self.subject_field.setGeometry(QtCore.QRect(80, 80, 701, 20))
        self.subject_field.setObjectName("subject_field")

        self.files = QtWidgets.QLineEdit(self.centralwidget)
        self.files.setGeometry(QtCore.QRect(80, 110, 701, 20))
        self.files.setObjectName("files")
        self.files.setReadOnly(True)

        self.to_whom = QtWidgets.QLineEdit(self.centralwidget)
        self.to_whom.setGeometry(QtCore.QRect(80, 50, 701, 20))
        self.to_whom.setObjectName("to_whom")
        self.to_whom.setReadOnly(True)

        self.from_whom = QtWidgets.QLineEdit(self.centralwidget)
        self.from_whom.setGeometry(QtCore.QRect(80, 20, 701, 20))
        self.from_whom.setObjectName("from_whom")
        self.from_whom.setReadOnly(True)

        self.given_sender = QtWidgets.QPushButton(self.centralwidget)
        self.given_sender.setGeometry(QtCore.QRect(20, 20, 51, 23))
        self.given_sender.setObjectName("given_sender")
        self.given_sender.clicked.connect(lambda: self.select_given_sender())

        self.list_e_mail_addresses = QtWidgets.QPushButton(self.centralwidget)
        self.list_e_mail_addresses.setGeometry(QtCore.QRect(20, 50, 51, 23))
        self.list_e_mail_addresses.setObjectName("list_e_mail_addresses")
        self.list_e_mail_addresses.clicked.connect(lambda: self.select_list_e_mail_addresses())

        self.directory_files = QtWidgets.QPushButton(self.centralwidget)
        self.directory_files.setGeometry(QtCore.QRect(20, 110, 51, 23))
        self.directory_files.setObjectName("directory_files")
        self.directory_files.clicked.connect(lambda: self.select_directory_files())

        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(10, 150, 781, 301))
        self.frame.setStyleSheet("background-color: rgb(200, 200, 200);")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")

        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(10, 10, 781, 131))
        self.frame_2.setStyleSheet("background-color: rgb(200, 200, 200);")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")

        self.frame.raise_()
        self.frame_2.raise_()
        self.topic.raise_()
        self.message.raise_()
        self.message_field.raise_()
        self.newsletter.raise_()
        self.subject_field.raise_()
        self.files.raise_()
        self.to_whom.raise_()
        self.from_whom.raise_()
        self.given_sender.raise_()
        self.list_e_mail_addresses.raise_()
        self.directory_files.raise_()

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def select_directory_files(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setFilter(QDir.Files)
        if dialog.exec_():
            self.files.setText(dialog.selectedFiles()[0]+"/")

    def select_list_e_mail_addresses(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.AnyFile)
        dialog.setNameFilter("Excel (*.xlsx)")
        if dialog.exec_():
            self.to_whom.setText(dialog.selectedFiles()[0])

    def select_given_sender(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.AnyFile)
        dialog.setNameFilter("Excel (*.xlsx)")
        if dialog.exec_():
            self.from_whom.setText(dialog.selectedFiles()[0])

    def start_newsletter(self):
        return

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Рассылка"))
        self.topic.setText(_translate("MainWindow", "Тема"))
        self.message.setText(_translate("MainWindow", "Сообщение"))
        self.newsletter.setText(_translate("MainWindow", "Запустить"))
        self.given_sender.setText(_translate("MainWindow", "От кого"))
        self.list_e_mail_addresses.setText(_translate("MainWindow", "Кому"))
        self.directory_files.setText(_translate("MainWindow", "Файлы"))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
