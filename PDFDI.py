# -*- coding: utf-8 -*-
"""
PDF DI ver 0.2.1
@author: Deep.I Inc. @Jongwon Kim
Revision date: 2021-01-11
See here for more information :
    https://deep-eye.tistory.com
    https://deep-i.net
"""
import os
import sys
import webbrowser

from PIL import Image
from PyPDF2 import PdfFileMerger,PdfFileReader
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog,QTableWidgetItem,QMessageBox,QDialog,QPushButton
from PyQt5 import uic, QtWidgets


FROM_CLASS = uic.loadUiType("main.ui")[0]
VER_CLASS = uic.loadUiType("version.ui")[0]

class PDFDI(QMainWindow,FROM_CLASS):
    def __init__(self):
        super().__init__()
        # UI load
        self.setupUi(self)
        self.version = 'PDF DI ver '+ '0.2.1'
        self.setWindowTitle(self.version)
        # button
        self.but_add.clicked.connect(self.searchFile)
        self.but_del.clicked.connect(self.removeTable)
        self.but_refresh.clicked.connect(self.refreshTable)
        self.but_up.clicked.connect(self.upperRow)
        self.but_down.clicked.connect(self.bellowRow)
        self.but_fusion.clicked.connect(self.mergePDF)
        self.but_ver.clicked.connect(self.versions)
        # table
        self.fname = []
        header = self.table.horizontalHeader()       
        header.setSectionResizeMode(0,  QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1,  QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2,  QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3,  QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4,  QtWidgets.QHeaderView.ResizeToContents)
        # FORMAT
        self.format = ['jpg','png','jpeg','JPG','PNG','JPEG']
    def versions(self):
        self.versionDialog = VERSION(self)
    def upperRow(self):
        try:
            idx = self.table.currentRow()
            if idx == 0 : return
            temp = self.fname[idx]
            del self.fname[idx]
            self.fname = self.fname[:idx-1] + [temp] +  self.fname[idx-1:] 
            self.table.selectRow(idx-1)
            self.initTable()
        except:pass
    def bellowRow(self):
        try:
            idx = self.table.currentRow()
            if idx == len(self.fname)-1 : return
            temp = self.fname[idx]
            del self.fname[idx]
            self.fname = self.fname[:idx+1] + [temp] +  self.fname[idx+1:] 
            self.table.selectRow(idx+1)
            self.initTable()
        except:pass
    def searchFile(self):
        fname = QFileDialog.getOpenFileNames(self, 
                                             '파일을 선택해주세요.', './',
                                             "PDF file (*.pdf) ;; Image file (*.jpg;*.png;*.jpeg)")
        if fname[0] == []: return
        else:
            self.fname = self.fname + fname[0]
            self.initTable()
    def removeTable(self):
        if self.fname == []:return
        del self.fname[self.table.currentRow()]
        self.initTable()
    def refreshTable(self):
        if self.fname == []:return
        self.fname = []
        self.initTable()
    def initTable(self):
        self.table.setRowCount(len(self.fname))
        for i,file in enumerate(self.fname):
            try:
                self.format.index(file.split('.')[-1])
                num = 1
                self.table.setItem(i,4, QTableWidgetItem(file.split('.')[-1]))
            except:
                pdf_reader = PdfFileReader(file)
                num = pdf_reader.numPages
                self.table.setItem(i,4, QTableWidgetItem("pdf"))

            size = os.path.getsize(file)
            name = file.split('/')[-1]
            path = '/'.join(file.split('/')[:-1])
            self.table.setItem(i,0, QTableWidgetItem(name))
            self.table.setItem(i,1, QTableWidgetItem(path))
            self.table.setItem(i,2, QTableWidgetItem(str(num)+' 페이지'))
            self.table.setItem(i,3, QTableWidgetItem("%.2f MB" % (size / (1024.0 * 1024.0)) ))
        self.table.update()
    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'PDF DI', '프로그램을 종료하시겠습니까?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.deleteLater()
            app.quit()
        else: event.ignore()
    def mergePDF(self):
        if self.fname == [] : 
            self.popUp(1)
            return
        sName = QFileDialog.getSaveFileName(self, self.tr("병합된 PDF 저장 경로와 이름을 지정해주세요"), "./", 
                                            self.tr("PDF File (*.pdf)"))
        if sName[0] == '': return
        else:
            self.sName = sName[0]
            self.PDF2PDFs()
    def PDF2PDFs(self):
        imformat = self.format
        img2pdf = []
        p = len(self.fname)
        self.status.setText('파일 준비중')
        try:
            pdf_merger = PdfFileMerger(strict=False)
            for j,file in enumerate(self.fname):
                self.progressBar.setValue(int((100*(j+1))/p)-1)
                try:
                    try:
                        self.status.setText('이미지 파일 병합중')
                        a = self.format.index(file.split('.')[-1])
                        image1 = Image.open(file)
                        im1 = image1.convert('RGB')
                        cfile = file.replace('.'+imformat[a],'.pdf')
                        img2pdf.append(cfile)
                        im1.save(cfile)
                        im1.close()
                        pdf_merger.append(cfile)
                    except: 
                        self.status.setText('PDF 파일 병합중')
                        pdf_merger.append(file)
                except:
                    self.status.setText('병합 실패')
                    self.popUp(0)
                    return
            self.status.setText('최종 파일 생성중')
            pdf_merger.write(self.sName)
            pdf_merger.close()
            for i in img2pdf:
                os.remove(i)
            self.status.setText('병합 성공')
            self.progressBar.setValue(100)
            self.mergePopUp()
        except : 
            self.status.setText('병합 실패')
            self.popUp(0)
            for i in img2pdf:
                os.remove(i)
    def popUp(self,idx):
        hmsg = ["PDF 병합 오류","PDF 병합 오류"]
        msg = ["PDF 병합이 실패했습니다. 파일을 확인해주세요.",
               "1개 이상의 PDF 파일을 지정해주세요."]
        QtWidgets.QMessageBox.information(self, hmsg[idx], msg[idx])
    def mergePopUp(self):
        msgbox = QMessageBox(self)
        msgbox.setWindowTitle("PDF 병합 성공")
        msgbox.setIcon(QMessageBox.Information)
        msgbox.setText("PDF 병합을 성공했습니다.")
        botonyes = QPushButton("경로 열기")
        msgbox.addButton(botonyes, QMessageBox.YesRole)
        botonno = QPushButton("확인")
        msgbox.exec_()
        if msgbox.clickedButton() == botonyes:
            os.startfile( os.path.dirname(self.sName))
        elif msgbox.clickedButton() == botonno:
            msgbox.deleteLater()
class VERSION(QDialog,VER_CLASS):
    def __init__(self,perent):
        super(VERSION,self).__init__(perent)
        self.setupUi(self)
        self.show()
        self.version = 'PDF DI ver '+ '0.2.1'
        self.setWindowTitle(self.version)
        self.but_close.clicked.connect(self.deleteLater)
        self.but_git.clicked.connect(lambda:webbrowser.open('https://github.com/DEEPI-LAB'))
        self.but_tis.clicked.connect(lambda:webbrowser.open('https://deep-eye.tistory.com/category/Program'))
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ShowApp = PDFDI()
    ShowApp.show()
    sys.exit(app.exec_())