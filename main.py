import io
import sqlite3
import sys
from datetime import date
from random import randrange
from openpyxl import load_workbook

from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QWidget, QMainWindow, QTableWidgetItem

class AddScreen(QMainWindow):
    def __init__(self, parent, summ):
        super().__init__(parent)
        self.parent = parent
        self.summ = summ
        self.setFixedSize(832, 629)
        uic.loadUi('add.ui', self)
        self.setWindowTitle('Добавление')
        now_d = date.today()
        self.dateEdit.setDate(now_d)
        self.doxod.setMaximum(100000000)
        self.doxod.setMinimum(0)
        self.rasxod.setMaximum(100000000)
        self.rasxod.setMinimum(0)
        self.addbtn.clicked.connect(self.add_func)

    def add_func(self):
        self.error.setText('')
        data = (str(self.dateEdit.date().toString('dd-MM-yyyy'))).replace('-', '/')
        doxod = self.doxod.value()
        rosxod = self.rasxod.value()
        state = self.state.text()
        if (self.summ - rosxod) < 0 or state == '':
            self.error.setText('Некорректные данные')
        else:
            self.parent.data.append((data, doxod, rosxod, self.summ - rosxod + doxod, state))
            self.parent.update_result(0)
            self.close()

class DeleteScreen(QMainWindow):
    def __init__(self, parent, summ):
        super().__init__(parent)
        self.parent = parent
        self.summ = summ
        self.setFixedSize(832, 629)
        uic.loadUi('delete.ui', self)
        self.setWindowTitle('Удаление')
        now_d = date.today()
        self.dateEdit.setDate(now_d)
        self.doxod.setMaximum(100000000)
        self.doxod.setMinimum(0)
        self.rasxod.setMaximum(100000000)
        self.rasxod.setMinimum(0)
        self.addbtn.clicked.connect(self.delete_func)

    def delete_func(self):
        self.error.setText('')
        data = (str(self.dateEdit.date().toString('dd-MM-yyyy'))).replace('-', '/')
        doxod = self.doxod.value()
        rosxod = self.rasxod.value()
        state = self.state.text()
        i = 0
        while i < len(self.parent.data):
            if data in self.parent.data[i] and doxod in self.parent.data[i] and rosxod in self.parent.data[i] and state in self.parent.data[i]:
                self.parent.data = self.parent.data[:i] + self.parent.data[i + 1:]
                break
            else:
                i += 1
        else:
            self.error.setText('Не найдено')
        self.parent.update_result(1)
        self.close()


class MainScreen(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setFixedSize(1085, 831)
        uic.loadUi('main.ui', self)
        self.setWindowTitle('Главная')
        self.calendarWidget.hide()
        self.dependtable.hide()
        self.find_button.hide()
        self.tableWidget.show()
        self.wb = load_workbook('data.xlsx')
        self.worksheet = self.wb.active
        self.data = []
        for i in range(len(self.worksheet['A'])):
            if self.worksheet['A'][i].value:
                self.data.append((self.worksheet['A'][i].value, self.worksheet['B'][i].value, self.worksheet['C'][i].value, self.worksheet['D'][i].value, self.worksheet['E'][i].value))
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setHorizontalHeaderLabels(['дата', 'сумма дохода', 'сумма расхода', 'оставшаяся сумма', 'статья дохода/расхода'])
        self.tableWidget.setRowCount(0)
        self.data = self.data[1:]
        for i in range(len(self.data)):
            if self.data[i][0]:
                self.tableWidget.setRowCount(
                    self.tableWidget.rowCount() + 1)
                self.tableWidget.setItem(i, 0, QTableWidgetItem(str(self.data[i][0])))
                self.tableWidget.setItem(i, 1, QTableWidgetItem(str(self.data[i][1])))
                self.tableWidget.setItem(i, 2, QTableWidgetItem(str(self.data[i][2])))
                self.tableWidget.setItem(i, 3, QTableWidgetItem(str(self.data[i][3])))
                self.tableWidget.setItem(i, 4, QTableWidgetItem(str(self.data[i][4])))
        self.search.textChanged.connect(self.search_func)
        if len(self.data) != 0:
            self.summ = self.data[-1][-2]
            self.money.setText(str(self.summ))
        else:
            self.summ = 0
            self.money.setText(str(self.summ))
        self.add.clicked.connect(self.add_func)
        self.dell.clicked.connect(self.delete_func)
        self.calendar.clicked.connect(self.calendar_func)
        self.glav.clicked.connect(self.all_func)

    def all_func(self):
        self.calendarWidget.hide()
        self.dependtable.hide()
        self.find_button.hide()
        self.tableWidget.show()

    def update_result(self, type):
        last_row = len(self.data) + 1
        val = 'A2:E' + str(last_row + type)
        for row in self.worksheet[val]:
            for cell in row:
                cell.value = None
        val = 'A2:E' + str(last_row)
        for row in range(len(self.worksheet[val])):
            for i in range(5):
                self.worksheet[val][row][i].value = self.data[row][i]
        self.wb.save('data.xlsx')
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setHorizontalHeaderLabels(
            ['дата', 'сумма дохода', 'сумма расхода', 'оставшаяся сумма', 'статья дохода/расхода'])
        self.tableWidget.setRowCount(0)
        for i in range(len(self.data)):
            self.tableWidget.setRowCount(
                self.tableWidget.rowCount() + 1)
            self.tableWidget.setItem(i, 0, QTableWidgetItem(str(self.data[i][0])))
            self.tableWidget.setItem(i, 1, QTableWidgetItem(str(self.data[i][1])))
            self.tableWidget.setItem(i, 2, QTableWidgetItem(str(self.data[i][2])))
            self.tableWidget.setItem(i, 3, QTableWidgetItem(str(self.data[i][3])))
            self.tableWidget.setItem(i, 4, QTableWidgetItem(str(self.data[i][4])))
        if len(self.data) != 0:
            self.summ = self.data[-1][-2]
            self.money.setText(str(self.summ))
        else:
            self.summ = 0
            self.money.setText(str(self.summ))

    def calendar_func(self):
        self.calendarWidget.show()
        self.dependtable.show()
        self.find_button.show()
        self.tableWidget.hide()
        self.find_button.clicked.connect(self.find_func)

    def find_func(self):
        data = self.calendarWidget.selectedDate().toString('dd-MM-yyyy').replace('-', '/')
        new_data = []
        for element in self.data:
            if data in element:
                new_data.append(element)
        self.dependtable.setColumnCount(5)
        self.dependtable.setHorizontalHeaderLabels(
            ['дата', 'сумма дохода', 'сумма расхода', 'оставшаяся сумма', 'статья дохода/расхода'])
        self.dependtable.setRowCount(0)
        for i in range(len(new_data)):
            self.dependtable.setRowCount(
                self.dependtable.rowCount() + 1)
            self.dependtable.setItem(i, 0, QTableWidgetItem(str(new_data[i][0])))
            self.dependtable.setItem(i, 1, QTableWidgetItem(str(new_data[i][1])))
            self.dependtable.setItem(i, 2, QTableWidgetItem(str(new_data[i][2])))
            self.dependtable.setItem(i, 3, QTableWidgetItem(str(new_data[i][3])))
            self.dependtable.setItem(i, 4, QTableWidgetItem(str(new_data[i][4])))

    def add_func(self):
        self.calendarWidget.hide()
        self.dependtable.hide()
        self.find_button.hide()
        self.tableWidget.show()
        self.add_form = AddScreen(self, self.summ)
        self.add_form.show()

    def delete_func(self):
        self.calendarWidget.hide()
        self.dependtable.hide()
        self.find_button.hide()
        self.tableWidget.show()
        self.delete_form = DeleteScreen(self, self.summ)
        self.delete_form.show()

    def search_func(self):
        self.calendarWidget.hide()
        self.dependtable.hide()
        self.find_button.hide()
        self.tableWidget.show()
        text = self.search.text()
        new_data = []
        for elements in self.data:
            if text in elements[0]:
                new_data.append(elements)
        if new_data:
            self.tableWidget.setColumnCount(5)
            self.tableWidget.setHorizontalHeaderLabels(
                ['дата', 'сумма дохода', 'сумма расхода', 'оставшаяся сумма', 'статья дохода/расхода'])
            self.tableWidget.setRowCount(0)
            for i in range(len(new_data)):
                self.tableWidget.setRowCount(
                    self.tableWidget.rowCount() + 1)
                self.tableWidget.setItem(i, 0, QTableWidgetItem(str(new_data[i][0])))
                self.tableWidget.setItem(i, 1, QTableWidgetItem(str(new_data[i][1])))
                self.tableWidget.setItem(i, 2, QTableWidgetItem(str(new_data[i][2])))
                self.tableWidget.setItem(i, 3, QTableWidgetItem(str(new_data[i][3])))
                self.tableWidget.setItem(i, 4, QTableWidgetItem(str(new_data[i][4])))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainScreen()
    ex.show()
    sys.exit(app.exec())
