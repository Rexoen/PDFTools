#!/usr/bin/python

import sys	
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtWidgets import QApplication , QMainWindow, QFileDialog, QMessageBox,QDialog
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
from PIL.ImageQt import ImageQt
from PDFTools import Ui_MainWindow
from PyQt5 import QtCore
from PyQt5.QtGui import QPainter
from DSPDialog import Ui_dsp_Dialog
from PyPDF2 import PdfFileReader,PdfFileWriter
import tempfile
from pdf2image import convert_from_path

def mergePdf(files,outFile):
    pdfFileWriter = PdfFileWriter()
    for file in files:
        pdfFileReader = PdfFileReader(file)
        numPages = pdfFileReader.getNumPages()
        for index in range(0,numPages):
            pageObj = pdfFileReader.getPage(index)
            pdfFileWriter.addPage(pageObj)
    # 添加完每页，再一起保存至文件中
    pdfFileWriter.write(open(outFile, 'wb'))

def splitPdf(readFile, outFile, pages):
    if readFile != '':
        pdfFileWriter = PdfFileWriter()
        # 获取 PdfFileReader 对象
        pdfFileReader = PdfFileReader(readFile)  # 或者这个方式：pdfFileReader = PdfFileReader(open(readFile, 'rb'))
        # 文档总页数
        numPages = pdfFileReader.getNumPages()
        #start_page -= 1 #把人类感官的页数（1开始）改成计算机认的页数（0开始）
        pages = pages.replace(' ','')
        for page in pages.split(','):
            if '-' not in page:
                pageObj = pdfFileReader.getPage(int(page)-1)
                pdfFileWriter.addPage(pageObj)
                print(int(page)-1)
            if '-' in page:
                start_page = page.split('-')[0]
                end_page = page.split('-')[1]
                for index in range(int(start_page)-1, int(end_page)):
                    pageObj = pdfFileReader.getPage(index)
                    pdfFileWriter.addPage(pageObj)
                    print(index)
        # 添加完每页，再一起保存至文件中
        pdfFileWriter.write(open(outFile, 'wb'))

def 双面打印(readFile):
    奇数页写 = PdfFileWriter()
    偶数页写 = PdfFileWriter()
    # 获取 PdfFileReader 对象
    读取的pdf内容 = PdfFileReader(readFile)  # 或者这个方式：pdfFileReader = PdfFileReader(open(readFile, 'rb'))
    # 文档总页数
    numPages = 读取的pdf内容.getNumPages()
    for index in range(0,numPages):
        if index % 2 == 0:   #奇数页
            pageObj = 读取的pdf内容.getPage(index)
            奇数页写.addPage(pageObj)
    # 添加完每页，再一起保存至文件中
    奇数页写.write(open('1.pdf', 'wb'))
   
    if numPages % 2 != 0: #判断文档总页码数奇偶性，来确定最后一页是否为偶数页
        numPages -= 1
        偶数页写.addBlankPage(读取的pdf内容.getPage(0)['/MediaBox'][2],读取的pdf内容.getPage(0)['/MediaBox'][3])
        #print(读取的pdf内容.getPage(0)['/MediaBox'][2],读取的pdf内容.getPage(0)['/MediaBox'][3])
        
    for index in range(numPages,-1,-1):  #偶数页倒序
        if index % 2 != 0:   #偶数页
            pageObj = 读取的pdf内容.getPage(index)
            偶数页写.addPage(pageObj)
    偶数页写.write(open('2.pdf', 'wb'))
    print('双面打印元文件已生成。')

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):    
        super(MainWindow, self).__init__()
        self.setupUi(self)
        self.np_button.clicked.connect(self.np_button_handler)
        self.dsp_button.clicked.connect(self.onDspPrint)
        self.extract_file_open_button.clicked.connect(self.extract_button_handler)
        self.extract_confirm_button.clicked.connect(self.extract_confirm_button_handler)
        self.add_item_button.clicked.connect(self.add_item_button_handler)
        self.merge_confirm_button.clicked.connect(self.merge_confirm_button_handler)
        self.clean_button.clicked.connect(self.clean_listView)
        self.litem = []
        #一定要在主窗口类的初始化函数中对子窗口进行实例化，如果在其他函数中实例化子窗口可能会出现子窗口闪退的问题
        self.DspDialog = DSPChildWindow()


    def spliPDF(self, pages):
        if '.pdf' in self.extract_label.text():
            splitPdf(readFile = self.extract_label.text(),outFile = self.extract_label.text().replace('.pdf','_splited.pdf'), pages =  pages)
            self.extract_lineEdit.setText("")
        else:
            msg_box = QMessageBox(QMessageBox.Warning, '警告', '请确定是否选定PDF文件')
            msg_box.exec_()

    
    def add_item_button_handler(self):
        fileopd = self.open_dialog_box()
        print(fileopd)
        self.litem.append(fileopd)
        qlm = QtCore.QStringListModel()
        qlm.setStringList(self.litem)
        self.merge_listView.setModel(qlm)
    
    def merge_confirm_button_handler(self):
        mergePdf(self.litem,'Merged.pdf')
        self.litem = []
        qlm = QtCore.QStringListModel()
        qlm.setStringList(self.litem)
        self.merge_listView.setModel(qlm)
        reply = QMessageBox.information(self, '提示','合并完成！',QMessageBox.Yes)
        
    def clean_listView(self):
        self.litem = []
        qlm = QtCore.QStringListModel()
        qlm.setStringList(self.litem)
        self.merge_listView.setModel(qlm)

    def extract_button_handler(self):
        self.extract_label.setText(self.open_dialog_box())
        print(self.extract_label.text())

    def extract_confirm_button_handler(self):
        self.spliPDF(self.extract_lineEdit.text())
        reply = QMessageBox.information(self, '提示','抽取完成！',QMessageBox.Yes)

    def open_dialog_box(self):
        fileDialog = QFileDialog(self)
        filename =  fileDialog.getOpenFileName(self, 'Open file', '', 'PDF (*.pdf)')

        path = filename[0]
        if '.pdf' not in path and path != '':
            print('%s:非预期文件类型' % path)
            msg_box = QMessageBox(QMessageBox.Warning, '警告', '请先选择PDF文件')
            msg_box.exec_()
        else:
            return path

    def np_button_handler(self):
        fileopd = self.open_dialog_box()
        if '*.pdf' in fileopd:
            self.printDialog(fileopd)

    def printDialog(self,filePath):
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec_() == QPrintDialog.Accepted:
            with tempfile.TemporaryDirectory() as path:
                images = convert_from_path(filePath, dpi=300, output_folder=path)
                painter = QPainter()
                painter.begin(printer)
                for i, image in enumerate(images):
                    if i > 0:
                        printer.newPage()
                    rect = painter.viewport()
                    qtImage = ImageQt(image)
                    qtImageScaled = qtImage.scaled(rect.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    painter.drawImage(rect, qtImageScaled)
                painter.end()
    
    def onDspPrint(self):
        双面打印(self.open_dialog_box())
        self.DspDialog.show()


class DSPChildWindow(QDialog, Ui_dsp_Dialog):
    def __init__(self):
        super(DSPChildWindow, self).__init__()
        self.setupUi(self)
        self.litem = []
        self.former_button.clicked.connect(self.onFBClick)
        self.latter_button.clicked.connect(self.onLBClick)

    def onFBClick(self):
        self.printDialog('./1.pdf')
        self.former_button.hide
    
    def onLBClick(self):
        self.printDialog('./2.pdf')
        self.latter_button.hide

    def printDialog(self,filePath):
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec_() == QPrintDialog.Accepted:
            with tempfile.TemporaryDirectory() as path:
                images = convert_from_path(filePath, dpi=300, output_folder=path)
                painter = QPainter()
                painter.begin(printer)
                for i, image in enumerate(images):
                    if i > 0:
                        printer.newPage()
                    rect = painter.viewport()
                    qtImage = ImageQt(image)
                    qtImageScaled = qtImage.scaled(rect.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    painter.drawImage(rect, qtImageScaled)
                painter.end()

if __name__=="__main__":  
    app = QApplication(sys.argv)  
    main_window = MainWindow()  
    dsp_chind_window = DSPChildWindow()
    main_window.show()
    sys.exit(app.exec_())  
