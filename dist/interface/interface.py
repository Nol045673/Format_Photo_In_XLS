from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QFileDialog, QApplication, QLabel
from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton
import sys
import os
import xlsxwriter
from PIL import Image
import getpass


class Window1(QtWidgets.QMainWindow):
    def __init__(self):
        super(Window1, self).__init__()
        loadUi('FirstWindowForXlsxFormat.ui', self)
        self.setWindowTitle('Window1')


class Hold_before(QWidget):
    def __init__(self):
        super(Hold_before, self).__init__()
        loadUi('SecondWindowForXlsxFormat.ui', self)
        self.setWindowTitle('Window2')


class Hold_after(QWidget):
    def __init__(self):
        super(Hold_after, self).__init__()
        loadUi('ThreeWindowForXlsxFormat.ui', self)
        self.setWindowTitle('Window3')


class Hold_after_ones(QWidget):
    def __init__(self):
        super(Hold_after_ones, self).__init__()
        loadUi('FourWindowForXlsxFormat.ui', self)
        self.setWindowTitle('Window4')


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setWindowTitle('MainWindow')
        self.spis_photo = []
        self.spis_to_remove = []
        self.workbook = xlsxwriter.Workbook('Holds.xlsx')
        self.numb = 0
        self.name_after = 'photo_after1_'

    def show_window_1(self):
        self.w1 = Window1()
        self.w1.pushButton.clicked.connect(self.show_window_2)
        self.w1.pushButton_2.clicked.connect(self.show_window_3)
        self.w1.show()

    def show_window_2(self):
        self.w2 = Hold_before()
        self.w2.pushButton.clicked.connect(self.load_photos)
        self.w2.pushButton_2.clicked.connect(self.load_photos)
        self.w2.pushButton_3.clicked.connect(self.load_photos)
        self.w2.pushButton_4.clicked.connect(self.load_photos)
        self.w2.pushButton_5.clicked.connect(self.load_photos)
        self.w2.pushButton_6.clicked.connect(self.load_photos)
        self.w2.pushButton_7.clicked.connect(self.load_photos)
        self.w2.pushButton_8.clicked.connect(self.load_first_page)
        self.w2.show()

    def show_window_3(self):
        self.w3 = Hold_after()
        self.w3.pushButton.clicked.connect(self.show_window_4)
        self.w3.pushButton_2.clicked.connect(self.show_window_4)
        self.w3.pushButton_3.clicked.connect(self.show_window_4)
        self.w3.pushButton_4.clicked.connect(self.show_window_4)
        self.w3.pushButton_5.clicked.connect(self.show_window_4)
        self.w3.pushButton_6.clicked.connect(self.show_window_4)
        self.w3.pushButton_7.clicked.connect(self.show_window_4)
        self.w3.pushButton_8.clicked.connect(self.finish)
        self.w3.show()

    def show_window_4(self):
        self.w4 = Hold_after_ones()
        self.w4.pushButton.clicked.connect(self.load_photos)
        self.w4.pushButton_2.clicked.connect(self.load_photos)
        self.w4.pushButton_3.clicked.connect(self.load_photos)
        self.w4.pushButton_4.clicked.connect(self.load_photos)
        self.w4.pushButton_5.clicked.connect(self.load_photos)
        self.w4.pushButton_7.clicked.connect(self.load_photos)
        self.w4.pushButton_6.clicked.connect(self.load_after_numb)
        self.w4.show()

    def load_photos(self):
        dialogSelectFiles = QFileDialog()
        dialogSelectFiles.setFileMode(QFileDialog.ExistingFiles)
        # прочитать список файлов и сохранить его в data
        fname = dialogSelectFiles.getOpenFileNames(self, 'Open file', '/home/axa/Stuff')
        self.spis_photo.append(fname[0])

    def load_after_numb(self):
        self.numb += 1
        if self.numb == 8:
            self.numb = 0
        worksheet = self.workbook.add_worksheet(f'Hold №{self.numb} After')
        bg_color = self.workbook.add_format({'bg_color': '#CCFFFF'})
        merge_format = self.workbook.add_format({'align': 'center'})
        format = self.workbook.add_format()
        format.set_font_size(18)
        worksheet.merge_range(0, 0, 0, 8, 'Merged Cells', merge_format)
        worksheet.write('A1', f'HOLD №{self.numb} CONDITION AFTER', bg_color)
        row = 9
        row_for_holds = 6
        for x in range(len(self.spis_photo)):
            worksheet.merge_range(row_for_holds, 8, row_for_holds, 20, 'Merged Cells', merge_format)
            if x == 0:
                worksheet.write(row_for_holds, 8, f'HATCH COVERS,WHEEL RAIL', merge_format)
            if x == 1:
                worksheet.write(row_for_holds, 8, f'COAMINGS, UNDER HATCH BALCONY', merge_format)
            if x == 2:
                worksheet.write(row_for_holds, 8, f'LADDERS/CORRUGATED BULKHEAD / PIPE BRACKETS / PIPE GUARDS /FORWARD, AFT ', merge_format)
            if x == 3:
                worksheet.write(row_for_holds, 8, f'SHELL FRAMES/ HOPPERS/ PORTSIDE, STARBOARD SIDE ', merge_format)
            if x == 4:
                worksheet.write(row_for_holds, 8, f' TANK TOP ', merge_format)
            if x == 5:
                worksheet.write(row_for_holds, 8, f'GLOVE TEST', merge_format)
            column = 0
            tg = self.spis_photo[x]
            for i in range(len(tg)):
                img = Image.open(tg[i])
                im = img.resize((300, 250))
                im.save(f'{self.name_after}{x}_{i}.jpg')
                self.spis_to_remove.append(f'{self.name_after}{x}_{i}.jpg')
                worksheet.insert_image(row, column, f'{self.name_after}{x}_{i}.jpg')
                column += 5
            row_for_holds += 17
            row += 17
        self.spis_photo = []
        if self.numb == 1:
            self.name_after = 'photo_after2_'
            self.w3.checkBox.setChecked(True)
        if self.numb == 2:
            self.name_after = 'photo_after3_'
            self.w3.checkBox_2.setChecked(True)
        if self.numb == 3:
            self.name_after = 'photo_after4_'
            self.w3.checkBox_3.setChecked(True)
        if self.numb == 4:
            self.name_after = 'photo_after5_'
            self.w3.checkBox_4.setChecked(True)
        if self.numb == 5:
            self.name_after = 'photo_after6_'
            self.w3.checkBox_5.setChecked(True)
        if self.numb == 6:
            self.name_after = 'photo_after7_'
            self.w3.checkBox_6.setChecked(True)
        if self.numb == 7:
            self.w3.checkBox_7.setChecked(True)
        self.w4.close()

    def load_first_page(self):
        worksheet = self.workbook.add_worksheet('Holds_Before')
        bg_color = self.workbook.add_format({'bg_color': '#CCFFFF'})
        merge_format = self.workbook.add_format({'align': 'center'})
        format = self.workbook.add_format()
        format.set_font_size(18)
        worksheet.merge_range(0, 0, 0, 8, 'Merged Cells', merge_format)
        worksheet.write('A1', f'HOLDS №1 - {len(self.spis_photo)} CONDITION BEFORE', bg_color)
        shapka = Image.open('img.png')
        ban = shapka.resize((800, 75))
        ban.save('ban.jpg')
        worksheet.insert_image(2, 5, 'ban.jpg')

        row = 9
        row_for_holds = 6

        for x in range(len(self.spis_photo)):
            worksheet.merge_range(row_for_holds, 8, row_for_holds, 13, 'Merged Cells', merge_format)
            worksheet.write(row_for_holds, 8, f'HOLD №{x + 1}', merge_format)
            column = 0
            tg = self.spis_photo[x]
            for i in range(len(tg)):
                img = Image.open(tg[i])
                im = img.resize((300, 250))
                im.save(f'photo{x}_{i}.jpg')
                self.spis_to_remove.append(f'photo{x}_{i}.jpg')
                worksheet.insert_image(row, column, f'photo{x}_{i}.jpg')
                column += 5
            row_for_holds += 17
            row += 17
        self.spis_photo = []
        self.w2.close()

    def finish(self):
        self.workbook.close()
        path = os.path.abspath('Holds.xlsx')
        self.w3.lineEdit.setText(path)
        for x in self.spis_to_remove:
            os.remove(x)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show_window_1()
    sys.exit(app.exec_())
