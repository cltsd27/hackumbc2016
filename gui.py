import sys
from PyQt4 import QtCore, QtGui
from form import Ui_Wizard
from docx import Document 
 
class MyDialog(QtGui.QWizard):
    def __init__(self, parent=None):
        QtGui.QWidget.__init__(self, parent)
        self.ui = Ui_Wizard()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.makeResume)

    def makeResume(self):
        document = Document()

        document.add_heading(str(self.ui.name.text()),0)
        p = document.add_paragraph(
            str(self.ui.address.text())
            )
        p.add_run(' | ')
        p.add_run(
            str(self.ui.phoneNumber.text())
            )
        p.add_run(' | ')
        p.add_run(
            str(self.ui.email.text())
            )
        if str(self.ui.websites.toPlainText()) != "":
            p.add_run(' | ')
            p.add_run(str(self.ui.websites.toPlainText()))
        document.add_heading("Objective", 1)
        p2 = document.add_paragraph(
            str(self.ui.objective.toPlainText())
            )
        document.add_heading("Skills and Technologies", 1)
        document.add_paragraph(
            str(self.ui.strongSkills.toPlainText()), style="ListBullet"
        )
        document.add_paragraph(
            str(self.ui.strongTech.toPlainText()), style="ListBullet"
        )
        document.add_page_break()
        document.save('resume.docx')

if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    myapp = MyDialog()
    myapp.show()
    sys.exit(app.exec_())


