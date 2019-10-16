import sys
from PyQt5.QtWidgets import (QWidget, QPushButton, QLabel, QLineEdit, QInputDialog, QApplication, QMessageBox)
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtCore import QSize
from PyQt5.QtGui import QImage, QPalette, QBrush, QIcon, QPixmap
from PyQt5.QtCore import pyqtSlot
import time
import os
import csv, time
import docx
from rabbit import Rabbit
from xml.etree.cElementTree import XML
from pathlib import Path
import ntpath
from zipfile import ZipFile
from zipfile import ZIP_DEFLATED
import shutil
#import ZawGyi_2_UniCode_converter_By__KP

class UserApp(QWidget):

    UserInput = ''
    paragraphs = []
    my_file = ''
    my_file2 = ''
    path = ''
    base = ''
    WORD_NAMESPACE = ''
    PARA = ''
    TEXT = ''

    def __init__(self):
        QWidget.__init__(self)
        self.initGUI()

    def initGUI(self):
        self.setWindowTitle("ZawGyi to UniCode Converter")
        self.setGeometry(10,10,300,150)

        label = QLabel(self)
        pixmap = QPixmap('zawgyimyanmarconverter.png')
        label.setPixmap(pixmap)
        self.move(100, 50)
        self.resize(280, 250)

        btn1 = QPushButton("Browser", self)
        btn1.move(10, 132)
        btn1.clicked.connect(self.on_add_button_clicked)
        label1 = QLabel(self)
        label1.setText("Please kindly browse your files.\nCurrently we can only choose a single Docx file. Thanks")
        label1.move(10,105)

        btn2 = QPushButton("Start",self)
        btn2.move(10,211)
        btn2.clicked.connect(self.ConvFun)
        label2 = QLabel(self)
        label2.setText("Click here to convert the file from ZawGyi to UniCode")
        label2.move(10,196)

        self.show()

    def on_add_button_clicked(self):
        
        self.file_path = QFileDialog.getOpenFileName(self, 'Open File', './',
                                                         filter="All Files(*.*);;Text Files(*.txt)")
                
        self.Input = str(self.file_path)
        length = len(self.Input)-20
        self.UserInput = self.Input[2:length]
        print (self.UserInput)

    def ConvFun(self):
        print(self.UserInput)
        self.my_file = ntpath.basename(str(self.UserInput))
        #print(self.my_file)
        self.my_file2 = ntpath.dirname(str(self.UserInput))
        #print(self.my_file2)
        self.base = os.path.splitext(self.my_file)[0]
        #print(self.base)

        self.WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        self.PARA = self.WORD_NAMESPACE + 'p'
        self.TEXT = self.WORD_NAMESPACE + 't'
        
        self.path = ntpath.dirname(str(self.UserInput))
        
        f = open(self.path + '\\' + "replace.csv",'w+', encoding='utf8')
        
        self.Document = ZipFile(self.my_file2 + '\\' + self.my_file)
        self.xml_content = self.Document.read('word/document.xml')
        self.Document.close()

        self.tree = XML(self.xml_content)
        
        for self.paragraph in self.tree.getiterator(self.PARA):
            for self.node in self.paragraph.getiterator(self.TEXT):
                if self.node.text:
                    self.u_conv_word = str(Rabbit.zg2uni(self.node.text))
                    print(self.u_conv_word)
                    print(self.node.text)
                    f.write("{0},{1}\n".format(self.node.text,self.u_conv_word))
        f.close()

        os.rename(self.my_file2 + '\\' + self.my_file, self.my_file2 + '\\'+ self.base + '.zip')

        with ZipFile(self.my_file2 + '\\'+ self.base + str(".zip"), 'r') as zipObj:
            zipObj.extractall(self.my_file2+'\\' + self.base)
            print(self.my_file2+'\\' + self.base)

        self.counter = 0

        xml_filename_original = self.my_file2 + '\\' + self.base + '\\word\\' + "document.xml"
        xml_filename_new = xml_filename_original.replace(".xml", "_new.xml")
        find_and_replace_filename = self.my_file2 + '\\' + "replace.csv"
        xml_file_original = open(xml_filename_original, "r", encoding='utf-8')
        xml_file_new = open(xml_filename_new, "w", encoding='utf-8')
        with open(find_and_replace_filename, "r", encoding='utf-8') as f:
            for line in xml_file_original:
                f.seek(0)
                reader = csv.reader(f)
                next(reader, None)
                for row in reader:
                    line_old = line
                    line_new = line.replace(row[0], row[1])
                    if line_new != line_old:
                        line = line_new
                        self.counter += 1
                        print("Counter: " + str(self.counter))
                xml_file_new.write(line)

        xml_file_original.close()
        xml_file_new.close()

        shutil.copy(self.path + '\\' + self.base + '\\word\\' + "document_new.xml", self.path + '\\' + self.base + '\\word\\' + "document.xml")

        zf = ZipFile("%s.zip" % (self.my_file2 + '\\' + self.base), "w", ZIP_DEFLATED)
        abs_src = os.path.abspath(self.my_file2 + '\\' + self.base)

        for dirname, subdirs, files in os.walk(self.my_file2 + '\\' + self.base):
            for filename in files:
                absname = os.path.abspath(os.path.join(dirname, filename))
                arcname = absname[len(abs_src) + 1:]
                print ('zipping %s as %s' % (os.path.join(dirname, filename),
                                        arcname))
                zf.write(absname, arcname)
        zf.close()

        os.rename(self.my_file2 + '\\' + self.base + ".zip", self.my_file2 + '\\'+ self.base + '_Uni.docx')
        print("Done")

        os.remove(self.path + '\\' + "replace.csv")
        shutil.rmtree(self.path + '\\' + self.base)        
        
app = QApplication(sys.argv)
ex = UserApp()
sys.exit(app.exec_())
