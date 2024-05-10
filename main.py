import shutil
from win32com import client
import sys
import csv
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import (QComboBox, QUndoCommand, QFileDialog, QHBoxLayout, QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox, QDateEdit, QHeaderView)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QColor, QIcon
import pandas as pd
import os
from openpyxl import load_workbook, Workbook
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment

class EditDateCommand(QUndoCommand):
    def __init__(self, date_edit, old_date, new_date):
        super().__init__()
        self.date_edit = date_edit
        self.old_date = old_date
        self.new_date = new_date

    def redo(self):
        self.date_edit.setDate(self.new_date)

    def undo(self):
        self.date_edit.setDate(self.old_date)

class EditComboCommand(QUndoCommand):
    def __init__(self, combo_box, old_value, new_value):
        super().__init__()
        self.combo_box = combo_box
        self.old_value = old_value
        self.new_value = new_value

    def redo(self):
        self.combo_box.setCurrentText(self.new_value)

    def undo(self):
        self.combo_box.setCurrentText(self.old_value)


class WheelIgnoredComboBox(QComboBox):
    def wheelEvent(self, event):
        event.ignore()

class WheelIgnoredDateEdit(QDateEdit):
    def wheelEvent(self, event):
        event.ignore()
        
class SchemeSelectionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Выбрать исполнительную схему')
        self.layout = QtWidgets.QVBoxLayout(self)

        
        self.table = QtWidgets.QTableWidget(0, 3)
        self.layout.addWidget(self.table)
        self.table.setHorizontalHeaderLabels(['Наименование схемы', '№', 'Примечание'])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)
        self.setGeometry(100, 100, 800, 600)
        self.populate_table()
        self.table.resizeColumnsToContents()

        self.ok_button = QtWidgets.QPushButton('OK')
        self.ok_button.clicked.connect(self.accept_selection)
        self.cancel_button = QtWidgets.QPushButton('Отмена')
        self.cancel_button.clicked.connect(self.reject)

        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        self.layout.addLayout(button_layout)

    def populate_table(self):
        scheme_data = self.parent().get_scheme_data()
        self.table.setRowCount(len(scheme_data))
        for row_index, (name, number, note) in enumerate(scheme_data):
            self.table.setItem(row_index, 0, QtWidgets.QTableWidgetItem(name))
            self.table.setItem(row_index, 1, QtWidgets.QTableWidgetItem(number))
            self.table.setItem(row_index, 2, QtWidgets.QTableWidgetItem(note))

    def accept_selection(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Ошибка', 'Не выбраны строки для добавления.')
            return

        selected_items = []
        for index in selected_rows:
            row = index.row()
            scheme = self.table.item(row, 0).text()
            number = self.table.item(row, 1).text()
            selected_items.append(f"{scheme} - {number}")

        selected_row = self.parent().table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, 'Ошибка', 'Не выбрана строка в журнале.')
            return

        existing_item = self.parent().table.item(selected_row, 13)
        existing_text = existing_item.text() if existing_item else ""
        new_text = "; ".join(selected_items)
        final_text = existing_text + "; " + new_text if existing_text else new_text
        self.parent().table.setItem(selected_row, 13, QtWidgets.QTableWidgetItem(final_text))
        self.accept()


class AgreementSelectionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Выбрать согласование')
        self.layout = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableWidget(0, 1)  
        self.layout.addWidget(self.table)
        self.table.setHorizontalHeaderLabels(['Наименование документа'])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)

        self.populate_table()
        self.table.resizeColumnsToContents()
        self.setGeometry(100, 100, 800, 600)
        self.ok_button = QtWidgets.QPushButton('OK')
        self.ok_button.clicked.connect(self.accept_selection)
        self.cancel_button = QtWidgets.QPushButton('Отмена')
        self.cancel_button.clicked.connect(self.reject)

        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        self.layout.addLayout(button_layout)

    def populate_table(self):
        agreement_data = self.parent().get_agreement_data()
        self.table.setRowCount(len(agreement_data))
        for row_index, (name, ) in enumerate(agreement_data):
            self.table.setItem(row_index, 0, QtWidgets.QTableWidgetItem(name))

    def accept_selection(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Ошибка', 'Не выбраны строки для добавления.')
            return

        selected_items = []
        for index in selected_rows:
            row = index.row()
            document = self.table.item(row, 0).text()
            selected_items.append(f"{document}")

        selected_row = self.parent().table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, 'Ошибка', 'Не выбрана строка в журнале.')
            return

        existing_item = self.parent().table.item(selected_row, 10)
        existing_text = existing_item.text() if existing_item and existing_item.text() != "Отсутствуют" else ""
        new_text = "; ".join(selected_items)
        final_text = existing_text + "; " + new_text if existing_text else new_text
        self.parent().table.setItem(selected_row, 10, QtWidgets.QTableWidgetItem(final_text))
        self.accept()


class MTRSelectionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Выбрать МТР')
        self.layout = QtWidgets.QVBoxLayout(self)

        
        self.table = QtWidgets.QTableWidget(0, 3)  
        self.layout.addWidget(self.table)
        self.table.setHorizontalHeaderLabels(['Объект контроля', 'Сертификаты, паспорта и иные документы', 'Акты входного контроля'])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)  
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)  

        self.setGeometry(100, 100, 800, 600)
        self.table.setStyleSheet("QTableWidget::item:selected { background-color: #add8e6; }")

        self.populate_table()

        self.table.resizeColumnsToContents()
        self.ok_button = QtWidgets.QPushButton('OK')
        self.ok_button.clicked.connect(self.accept_selection)
        self.cancel_button = QtWidgets.QPushButton('Отмена')
        self.cancel_button.clicked.connect(self.reject)
        
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        self.layout.addLayout(button_layout)
    

    def populate_table(self):
        mtr_data = self.parent().get_mtr_data()
        self.table.setRowCount(len(mtr_data))

        for row_index, row_data in enumerate(mtr_data):
            for col_index in range(0, self.table.columnCount()):
                # Ensure that the column index does not exceed the length of row_data
                if col_index < len(row_data):
                    self.table.setItem(row_index, col_index, QtWidgets.QTableWidgetItem(row_data[col_index]))
                else:
                    # Optionally handle or log the missing data case
                    print(f"Missing data for row {row_index} column {col_index}")

    def accept_selection(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Ошибка', 'Не выбраны строки для добавления.')
            return

        selected_items = []
        for index in selected_rows:
            row = index.row()
            name = self.table.item(row, 0).text()
            qty = self.table.item(row, 1).text()
            unit = self.table.item(row, 2).text()
            selected_items.append(f"{name} - {qty} {unit}")


        
        selected_row = self.parent().table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, 'Ошибка', 'Не выбрана строка в журнале.')
            return

        existing_item = self.parent().table.item(selected_row, 9)
        existing_text = existing_item.text() if existing_item and existing_item.text() != "Не применялись" else ""
        new_text = "; ".join(selected_items)
        final_text = existing_text + "; " + new_text if existing_text else new_text
        self.parent().table.setItem(selected_row, 9, QtWidgets.QTableWidgetItem(final_text))
        self.accept()


class VolumeSelectionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Выбрать объём работы')
        self.layout = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableWidget(0, 3)  
        self.layout.addWidget(self.table)
        self.setGeometry(100, 100, 800, 600)
        self.ok_button = QtWidgets.QPushButton('OK')
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button = QtWidgets.QPushButton('Отмена')
        self.cancel_button.clicked.connect(self.reject)
        self.table.resizeColumnsToContents()
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        self.layout.addLayout(button_layout)

        self.selected_data = None

        self.table.cellDoubleClicked.connect(self.cell_double_clicked)

    def cell_double_clicked(self, row, column):
        
        num, ok = QtWidgets.QInputDialog.getInt(self, "Введите объем", "Объем работы:")
        if ok:
            
            selected_row = self.parent().table.currentRow()
            existing_data_item = self.parent().table.item(selected_row, 7)

            
            existing_data = existing_data_item.text() if existing_data_item else ""
            
            work_type = self.table.item(row, 0).text()
            volume = self.table.item(row, 1).text()
            unit = self.table.item(row, 2).text() if self.table.item(row, 2) else ""
            
            
            new_selection = f"{unit}_{work_type}_{num}_{volume}"

            
            if existing_data:
                self.selected_data = f"{existing_data}; {new_selection}".strip()
            else:
                self.selected_data = new_selection
        
            try:
                with open('docs/вор.csv', 'a+', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    file.seek(0)
                    reader = csv.reader(file)

                    _, work_type, num, volume = new_selection.split("_")
                    aosr_number = self.parent().table.item(self.parent().table.currentRow(), 0).text()
                    
                    writer.writerow([work_type, volume, num, aosr_number])
            except Exception as e:
                QMessageBox.warning(self, 'Ошибка', f'Не удалось сохранить данные: {e}')
            self.accept()
        else:
            self.selected_data = None
            self.reject()  

    def set_data(self, data):
        self.table.clear()
        self.table.setRowCount(len(data))
        self.table.setColumnCount(3)
        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                item = QtWidgets.QTableWidgetItem(value)
                self.table.setItem(row_index, col_index, item)
        self.table.resizeColumnsToContents()
        self.table.setColumnHidden(2, True)




class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return int(self.text()) < int(other.text())
        except ValueError:
            return self.text() < other.text()

class EditCellCommand(QUndoCommand):
    def __init__(self, table, row, column, old_value, new_value):
        super().__init__()
        self.table = table
        self.row = row
        self.column = column
        self.old_value = old_value
        self.new_value = new_value
        self.setText(f"Edit cell at ({row}, {column})")

    def redo(self):
        self.table.item(self.row, self.column).setText(self.new_value)

    def undo(self):
        self.table.item(self.row, self.column).setText(self.old_value)

class AOSRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.undo_stack = QtWidgets.QUndoStack(self)
        self.row_modified = {}
        self.initUI()
        

    def initUI(self):
        self.setWindowTitle('Журнал АОСР')
        self.setGeometry(100, 100, 800, 600)
        self.tab_widget = QTabWidget(self)
        self.tab_widget.setTabPosition(QTabWidget.South)
        self.setCentralWidget(self.tab_widget)

        self.tabs = {}
        tab_names = [
            ('Журнал АОСР', 'images/2.png'),
            ('Информация', 'images/7.png'),
            ('Ведомость МТР', 'images/1.png'),
            ('ВОР', 'images/1.png'),
            ('Исп. схемы', 'images/3.png'),
            ('Согласования', 'images/6.png'),
            ('Виды и объемы работ', 'images/5.png'),
            ('Реестр ИД', 'images/4.png'),
        ]
        for name, icon_path in tab_names:
            tab = QWidget()
            self.tabs[name] = tab
            self.tab_widget.addTab(tab, QIcon(icon_path), name)  

        self.setupJournalAOSRTab()
        self.setupOtherTabs()
        self.load_table_data()

        self.table.itemChanged.connect(self.item_changed)

    def setupJournalAOSRTab(self):
        layout = QVBoxLayout()
        self.tabs['Журнал АОСР'].setLayout(layout)
        self.table = QTableWidget(0, 14)  
        
        self.table.setHorizontalHeaderLabels([
            'Номер\nАОСР', 'Город\nстроительства', 'Наименование\n объекта', 'Дата\nначала\nработ', 'Дата\nокончания\nработ',
            'Дата подписания\nакта', '1. К\nосвидетельствованию\nпредъявлены\nработы', 'Вид и\nобъём\nработ\n(выгрузка ведомости)',
            'Работы выполнены\nпо\nпроектно-сметной\nдок.', 'При\nвыполнении работ\nприменены', 'При\nвыполнении работ\nотклонения',
            'Разрешается\nдля\nпроизводства работ', 'Сформировать\nакт', 'Исполнительная\nсхема'
        ])
        self.table.resizeRowsToContents()  
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setStretchLastSection(True)
        self.table.itemChanged.connect(self.capture_change)
        self.initialize_default_values()
        for column in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(column)
            header_item.setToolTip(header_item.text().replace('\n', ' '))  

        button_panel = QWidget()
        button_layout = QHBoxLayout()
        button_panel.setLayout(button_layout)
        self.table.setColumnHidden(1, True)
        self.table.setColumnHidden(2, True)
        
        btn_create_act = QPushButton('Создать акт')
        btn_create_act.clicked.connect(self.export_to_pdf_and_xls)
        btn_select_volume = QPushButton('Выбрать объём')
        btn_select_volume.clicked.connect(self.open_volume_selection)
        btn_select_mtr = QPushButton('Выбрать МТР')
        btn_select_mtr.clicked.connect(self.open_mtr_selection)
        btn_agreements = QPushButton('Выбрать согласования')
        btn_agreements.clicked.connect(self.open_agreement_selection)
        btn_select_scheme = QPushButton('Выбрать схему')
        btn_select_scheme.clicked.connect(self.open_scheme_selection)
        btn_clear_cell = QPushButton('Очистить ячейку')
        btn_clear_cell.clicked.connect(self.clear_selected_cell)
        upload_button = QPushButton(QIcon('images/import.jpeg'), '')
        upload_button.clicked.connect(self.upload_excel_data)

        download_button = QPushButton(QIcon('images/export.jpeg'), '')
        download_button.clicked.connect(self.download_csv_data)
        btn_undo = QPushButton(QIcon('images/reverse.jpeg'), '')
        btn_undo.clicked.connect(self.undo_stack.undo)
        
        new_button = QPushButton(QIcon('images/new.jpeg'), '')
        new_button.clicked.connect(self.new_project)

        button_layout.addWidget(new_button)
        button_layout.addWidget(btn_create_act)
        button_layout.addWidget(btn_clear_cell)
        button_layout.addWidget(btn_select_mtr)
        button_layout.addWidget(btn_select_volume)
        button_layout.addWidget(btn_agreements)
        button_layout.addWidget(btn_select_scheme)
        button_layout.addWidget(upload_button)
        button_layout.addWidget(download_button)
        button_layout.addWidget(btn_undo)
        layout.addWidget(button_panel)
        layout.addWidget(self.table)

        self.table.itemChanged.connect(self.item_changed)
        btn_add = QPushButton('Добавить запись')
        btn_remove = QPushButton('Удалить запись')
        btn_add.clicked.connect(self.add_record)
        btn_remove.clicked.connect(self.remove_record)
        layout.addWidget(btn_add)
        layout.addWidget(btn_remove)
    
    def new_project(self):
        try:
            # Step 1: Clear CSV files by writing empty data with headers only
            folder_path = 'docs/'
            files = os.listdir(folder_path)
            csv_files = [f for f in files if f.endswith('.csv')]
            for file_name in csv_files:
                with open(f'{folder_path}{file_name}', 'w', newline='', encoding='utf-8') as file:
                    pass

            # Step 2: Reset Application Tables
            for table_name, table in self.other_tables.items():
                table.setRowCount(0)  # Clear the table

            self.table.setRowCount(0)  # Clear the main table

            QMessageBox.information(self, 'Успех', 'Проект успешно создан.')
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Не удалось создать проект: {str(e)}')
    
    def capture_change(self, item):
        # On cell change, push an edit command to the undo stack
        old_value = item.data(QtCore.Qt.UserRole)  # UserRole used to store old value
        new_value = item.text()
        if item.column() == 7 and old_value != new_value:
            self.export_work_volume_to_general_ledger()
        if old_value is None or old_value != new_value:
            command = EditCellCommand(self.table, item.row(), item.column(), old_value, new_value)
            self.undo_stack.push(command)
        item.setData(QtCore.Qt.UserRole, new_value)  # Update the UserRole with new value



    def clear_selected_cell(self):
        selected_items = self.table.selectedItems()
        if selected_items:
            for item in selected_items:
                row, col = item.row(), item.column()
                # Verify the item still exists in the table
                if self.table.item(row, col) is not None:
                    self.table.item(row, col).setText('')  # Safely set text
                    if col == 7:
                        self.export_work_volume_to_general_ledger()

    
    def export_registry_to_xls_and_pdf(self):
        try:
            # Setup the paths
            xls_path = 'Acts/XLS/Реестр_ИД.xlsx'
            pdf_path = 'Acts/PDF/Реестр_ИД.pdf'
            os.makedirs('Acts/XLS', exist_ok=True)
            os.makedirs('Acts/PDF', exist_ok=True)

            # Check if the workbook already exists and clear existing data except the header
            if os.path.exists(xls_path):
                wb = openpyxl.load_workbook(xls_path)
                ws = wb.active
                ws.delete_rows(2, ws.max_row)  # Remove all rows except the header
            else:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(['№п/п', '№ акта', 'Наименование', 'Кол-во листов', 'Примечание'])  # Header row

            # Set worksheet title
            ws.title = "Реестр ИД"

            # Extract data from the table
            table = self.other_tables['Реестр ИД']
            row_count = 1  # Starting count for the rows after the header

            # Populate the worksheet with rows from the table
            for row_index in range(table.rowCount()):
                row_data = [table.item(row_index, col).text() if table.item(row_index, col) else '' for col in range(table.columnCount())]
                row_data.insert(0, str(row_count))  # Add row count at the beginning
                ws.append(row_data)
                row_count += 1

            # Set fixed column widths
            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 72
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 15

            ws['C1'].alignment = Alignment(horizontal='center')

            # Adjust the page layout settings to fit all columns on one page
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0

            # Save the workbook
            wb.save(xls_path)
            wb.close()

            # Convert XLS to PDF using Excel
            xlApp = client.Dispatch("Excel.Application")
            books = xlApp.Workbooks.Open(os.path.abspath(xls_path))
            ws = books.Worksheets[0]
            ws.Visible = 1
            ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            books.Close()
            xlApp.Quit()

            QMessageBox.information(self, 'Успех', 'Реестр ИД успешно экспортирован в XLS и PDF.')
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Не удалось экспортировать Реестр ИД: {str(e)}')
    
    def export_vor_to_xls_and_pdf(self):
        try:
            # Setup the paths
            xls_path = 'Acts/XLS/Ведомость_объёмов_работ.xlsx'
            pdf_path = 'Acts/PDF/Ведомость_объёмов_работ.pdf'
            os.makedirs('Acts/XLS', exist_ok=True)
            os.makedirs('Acts/PDF', exist_ok=True)

            # Check if the workbook already exists and clear existing data except the header
            if os.path.exists(xls_path):
                wb = openpyxl.load_workbook(xls_path)
                ws = wb.active
                ws.delete_rows(2, ws.max_row)  # Remove all rows except the header
            else:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(['№п/п', 'Наименование выполненных работ', 'Ед. изм', 'Кол-во', 'Примечание'])  # Header row

            # Set worksheet title
            ws.title = "ВОР"

            # Extract data from the table
            table = self.other_tables['ВОР']
            row_count = 1  # Starting count for the rows after the header

            # Populate the worksheet with rows from the table
            for row_index in range(table.rowCount()):
                row_data = [table.item(row_index, col).text() if table.item(row_index, col) else '' for col in range(table.columnCount())]
                row_data.insert(0, str(row_count))  # Add row count at the beginning
                ws.append(row_data)
                row_count += 1

            # Set fixed column widths
            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 72
            ws.column_dimensions['C'].width = 10
            ws.column_dimensions['D'].width = 10
            ws.column_dimensions['E'].width = 15

            ws['B1'].alignment = Alignment(horizontal='center')

            # Adjust the page layout settings to fit all columns on one page
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0

            # Save the workbook
            wb.save(xls_path)
            wb.close()

            # Convert XLS to PDF using Excel
            xlApp = client.Dispatch("Excel.Application")
            books = xlApp.Workbooks.Open(os.path.abspath(xls_path))
            ws = books.Worksheets[0]
            ws.Visible = 1
            ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            books.Close()
            xlApp.Quit()

            QMessageBox.information(self, 'Успех', 'ВОР успешно экспортирован в XLS и PDF.')
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Не удалось экспортировать ВОР: {str(e)}')




    def open_mtr_selection(self):
        dialog = MTRSelectionDialog(self)
        dialog.exec_()
    def setupOtherTabs(self):
        self.other_tables = {}
        tab_configs = {
            'Исп. схемы': ['Наименование схем', '№', ''],
            'Согласования': ['Наименование документа'],
            'Информация' : [''],
            'Виды и объемы работ': ['', 'Ед.изм', 'Номер'],
            'Реестр ИД': ['№ акта', 'Наименование', 'Кол-во листов', 'Примечание'],
            'Ведомость МТР': ['Объект контроля', 'Сертификаты, паспорта и иные документы', '*Акты входного контроля'],
            'ВОР': ['Наименование выполненных работ', 'Ед.изм', 'Кол-во', 'Примечание'],
        }
        for name, headers in tab_configs.items():
                tab = self.tabs[name]
                layout = QVBoxLayout()
                table = QTableWidget(0, len(headers))
                table.resizeRowsToContents()  
                table.setHorizontalHeaderLabels(headers)
                table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                table.horizontalHeader().setStretchLastSection(True)
                self.initialize_default_values()
                for column in range(table.columnCount()):
                    header_item = table.horizontalHeaderItem(column)
                    header_item.setToolTip(header_item.text())  

                tab.setLayout(layout)
                layout.addWidget(table)

                button_panel = QWidget()
                button_layout = QHBoxLayout()
                button_panel.setLayout(button_layout)

                btn_add = QPushButton('Добавить запись')
                btn_remove = QPushButton('Удалить запись')
                
                btn_add.clicked.connect(lambda checked, t=table: self.add_other_record(t))
                btn_remove.clicked.connect(lambda checked, t=table: self.remove_other_record(t))
                
                button_layout.addWidget(btn_add)
                button_layout.addWidget(btn_remove)

                if name == "Реестр ИД":
                    btn_refresh_reg = QPushButton('Обновить реестр')
                    btn_refresh_reg.clicked.connect(self.reload_reg)
                    button_layout.addWidget(btn_refresh_reg)
                    btn_export_reg = QPushButton('Экспорт реестра')
                    btn_export_reg.clicked.connect(self.export_registry_to_xls_and_pdf)
                    button_layout.addWidget(btn_export_reg)
                if name == 'ВОР':
                    btn_export_reg = QPushButton('Экспорт ведомости')
                    btn_export_reg.clicked.connect(self.export_vor_to_xls_and_pdf)
                    button_layout.addWidget(btn_export_reg)
                else:
                    btn_save = QPushButton('Сохранить записи')
                    button_layout.addWidget(btn_save)
                    btn_save.clicked.connect(lambda: self.save_all_tables())
                layout.addWidget(button_panel)

                self.other_tables[name] = table


    def add_other_record(self, table):
        row_count = table.rowCount()
        table.insertRow(row_count)
        for col_index in range(table.columnCount()):
            item = QTableWidgetItem('')
            table.setItem(row_count, col_index, item)

    def remove_other_record(self, table):
        current_row = table.currentRow()
        if current_row != -1:
            table.removeRow(current_row)
        else:
            QMessageBox.warning(self, 'Ошибка', 'Выберите строку для удаления')
    
    def open_agreement_selection(self):
        dialog = AgreementSelectionDialog(self)
        dialog.exec_()

    def open_scheme_selection(self):
        dialog = SchemeSelectionDialog(self)
        dialog.exec_()

    def initialize_default_values(self):
        for row in range(self.table.rowCount()):
            self.set_default_value(row)


    def upload_excel_data(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Открыть файл", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                wb = load_workbook(file_path, data_only=True)  # Ensure formulas are not read but their values
                for sheet_name in wb.sheetnames:
                    sheet = wb[sheet_name]
                    # Directly creating DataFrame from the Excel sheet
                    data = pd.DataFrame(sheet.values)
                    headers = data.iloc[0]  # Assuming the first row is the header
                    data = data[1:]  # Remove the header row from the data
                    data.columns = headers  # Set the header row as the DataFrame column names

                    csv_path = f"docs/{sheet_name}.csv"
                    data.to_csv(csv_path, index=False, header=True, encoding='utf-8')
                self.load_table_data()
                QMessageBox.information(self, 'Успех', 'Данные загружены.')
            except Exception as e:
                QMessageBox.warning(self, 'Ошибка', f'Не удалось загрузить данные: {str(e)}')

    def download_csv_data(self):
        folder_path = 'docs/'
        files = os.listdir(folder_path)
        csv_files = [f for f in files if f.endswith('.csv')]
        
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove the default sheet

        for file_name in csv_files:
            data = pd.read_csv(f'{folder_path}{file_name}')
            ws = wb.create_sheet(title=file_name.replace('.csv', ''))

            for col_num, col_name in enumerate(data.columns, start=1):
                ws.cell(row=1, column=col_num, value=col_name)
            
            for row_num, row in enumerate(data.itertuples(index=False, name=None), start=2):
                for col_num, item in enumerate(row, start=1):   
                    ws.cell(row=row_num, column=col_num, value=item)
        
        wb.save('docs/project.xlsx')
        QMessageBox.information(self, 'Успех', 'Все данные выгружены в Excel файл project.xlsx.')


    def set_default_value(self, row_index):
        item_8 = QTableWidgetItem('Не применялась')
        item_9 = QTableWidgetItem('Отсутствуют')
        combo_box_11 = WheelIgnoredComboBox()
        combo_box_11.addItems(['Да', 'Нет'])  
        combo_box_11.setCurrentText('Да')
        combo_box_11.setProperty('oldValue', combo_box_11.currentText())
        
        item_8.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        item_9.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)

        
        self.table.setItem(row_index, 9, item_8)
        self.table.setItem(row_index, 10, item_9)
        self.table.setCellWidget(row_index, 12, combo_box_11)
        combo_box_11.currentTextChanged.connect(lambda new_value, box=combo_box_11: self.update_registry_on_change(box, new_value))

    def update_registry_on_change(self, combo_box, new_value):
        with open('docs/реестр_ид.csv', 'w+', newline='', encoding='utf-8') as reg_file:
                reg_writer = csv.writer(reg_file, quoting=csv.QUOTE_ALL)
                for row in range(self.table.rowCount()):
                    if self.table.cellWidget(row, 12) and self.table.cellWidget(row, 12).currentText() == 'Да':
                        aosr_number = self.table.item(row, 0).text()
                        column_6_value = self.table.item(row, 6).text() if self.table.item(row, 6) else ''
                        column_13_value = self.table.item(row, 13).text().split('; ') if self.table.item(row, 13) else ''
                        reg_writer.writerow([f"АОСР № {aosr_number}", column_6_value, None, None])
                        for itemss in column_13_value:
                            reg_writer.writerow(['', itemss, None, None])
                        selected_mtrs = self.get_selected_mtr_data(row)
                        for mtr in selected_mtrs:
                            reg_writer.writerow(['', mtr[2], None, None])
        old_value = combo_box.property('oldValue') or combo_box.currentText()  # Initialize the old value if not set
        if old_value != new_value:
            command = EditComboCommand(combo_box, old_value, new_value)
            self.undo_stack.push(command)
        combo_box.setProperty('oldValue', new_value)  # Update the old value property

    def export_to_pdf_and_xls(self):
        self.save_changes()
        journal_path = 'docs/журнал_аоср.csv'
        info_path = 'docs/информация.csv'
        template_path = 'docs/blank.xlsx'
        registry_path = 'docs/реестр_ид.csv'
        xls_output_dir = 'Acts/XLS'
        pdf_output_dir = 'Acts/PDF'
        default_font = Font(name='Times New Roman', size=11)
        
        if os.path.exists(xls_output_dir):
            shutil.rmtree(xls_output_dir)
        os.makedirs(xls_output_dir, exist_ok=True)
        if os.path.exists(pdf_output_dir):
            shutil.rmtree(pdf_output_dir)
        os.makedirs(pdf_output_dir, exist_ok=True)
        
        
        try:
            journal_df = pd.read_csv(journal_path, header=None)
            info_df = pd.read_csv(info_path, header=None)
            
            with open(registry_path, 'w+', newline='', encoding='utf-8') as reg_file:
                reg_writer = csv.writer(reg_file, quoting=csv.QUOTE_ALL)
                
                for row in range(self.table.rowCount()):
                    if self.table.cellWidget(row, 12) and self.table.cellWidget(row, 12).currentText() == 'Да':
                        aosr_number = self.table.item(row, 0).text()
                        column_6_value = self.table.item(row, 6).text() if self.table.item(row, 6) else ''
                        column_13_value = self.table.item(row, 13).text().split('; ') if self.table.item(row, 13) else ''
                        reg_writer.writerow([f"АОСР № {aosr_number}", column_6_value, None, None])
                        for itemss in column_13_value:
                            reg_writer.writerow(['', itemss, None, None])
                        selected_mtrs = self.get_selected_mtr_data(row)
                        for mtr in selected_mtrs:
                            reg_writer.writerow(['', mtr[2], None, None])
                        
                        xls_filename = f"{xls_output_dir}/{aosr_number}.xlsx"
                        pdf_filename = f"{pdf_output_dir}/{aosr_number}.pdf"
                    
                    
                        with open(template_path, "rb") as f_template, open(xls_filename, "wb") as f_output:
                            f_output.write(f_template.read())
                        
                        wb = load_workbook(xls_filename)
                        sheet = wb.active

                        for row in sheet.iter_rows():
                            for cell in row:
                                cell.font = default_font

                        cell_mapping = ['F3', 'A6', 'A8', 'D41', 'D42', 'H6', 'A29', None, 'A32', 'A35', 'A39', 'A47', None, None]
                        journal_row_data = journal_df.iloc[row].tolist()
                        for index, cell in enumerate(cell_mapping):
                            if cell:
                                cell_value = sheet[cell]
                                current_font_size = cell_value.font.size - 1  # Уменьшение шрифта на 1 пункт
                                cell_value.font = Font(size=current_font_size)
                                if cell in ['H6', 'D41', 'D42']:
                                    date_str = journal_row_data[index]
                                    formatted_date = datetime.datetime.strptime(date_str, '%m/%d/%Y').strftime('%d.%m.%Y')
                                    sheet[cell] = formatted_date
                                elif cell == 'A35':
                                    new_str = journal_row_data[index].split("; ")
                                    ddata = ''
                                    for item in new_str:
                                        ddata += item
                                        ddata += '\n'
                                    sheet[cell] = ddata.strip()
                                    sheet.row_dimensions[sheet[cell].row].height = 15 * len(new_str)
                                else:
                                    sheet[cell] = journal_row_data[index]
                                sheet[cell].alignment = Alignment(wrapText=True)
                        
                        info_mapping = {'A13': 4, 'A16': 7, 'A19': 10, 'A23': 14, 'A25': 16, 'H51': (22, 1), 'H55': (26, 1), 'H58': (29, 1), 'H61': (32, 1)}
                        for cell, ref in info_mapping.items():
                            if isinstance(ref, tuple):
                                sheet[cell] = info_df.iloc[ref[0], ref[1]]
                            else:
                                sheet[cell] = info_df.iloc[ref, 0]
                        
                        wb.save(xls_filename)
                        wb.close()
                        
                        xlApp = client.Dispatch("Excel.Application")
                        books = xlApp.Workbooks.Open(os.path.abspath(xls_filename))
                        ws = books.Worksheets[0]
                        ws.Visible = 1
                        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_filename))
                        books.Close()
            
            self.reload_reg()

            QMessageBox.information(self, 'Успешно', f'Акты сформированы')
        except Exception as e:
           QMessageBox.warning(self, 'Ошибка', f'Не удалось сформировать акт {e}')
            
                
    def get_selected_mtr_data(self, aosr_row):
        mtr_table = self.other_tables['Ведомость МТР']
        selected_mtrs = []
        

        selected_ids = self.table.item(aosr_row, 9).text().split('; ') if self.table.item(aosr_row, 9) else []
        for mtr_id in selected_ids:
            for row in range(mtr_table.rowCount()):
                mtr_number = mtr_table.item(row, 1).text() if mtr_table.item(row, 1) else ''
                if mtr_number in mtr_id:
                    scheme = mtr_table.item(row, 0).text() if mtr_table.item(row, 0) else ''
                    quality_docs = mtr_table.item(row, 2).text() if mtr_table.item(row, 2) else ''
                    selected_mtrs.append((scheme, mtr_number, quality_docs))
        return selected_mtrs

    def create_xlsx_from_template(self, registry_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Registry Data"
        
        headers = ['АОСР №', 'Наименование объекта', 'Документы и сертификаты о качестве', 'Исполнительные схемы']
        ws.append(headers)
        
        with open(registry_path, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            for row in reader:
                ws.append(row)
        
        output_xlsx_path = 'docs/реестр_ид.xlsx'
        wb.save(output_xlsx_path)
        wb.close()
    
    def tab_changed(self, index):
        tab_name = self.tab_widget.tabText(index)
        self.setWindowTitle(f'АОСР - {tab_name}')

    def load_table_data(self):
        self.table.itemChanged.disconnect(self.item_changed)
        try:
            with open('docs/журнал_аоср.csv', 'r', encoding='utf-8') as file:
                self.table.setRowCount(0)
                reader = csv.reader(file)
                for row_index, row_data in enumerate(reader):
                    self.table.insertRow(row_index)
                    self.set_default_value(row_index)
                    for col_index, value in enumerate(row_data):
                        if "Unnamed" in value:
                            value = ""
                        if col_index in [3, 4, 5]:
                            date = QDate.fromString(value, "MM/dd/yyyy") if value else QDate.currentDate()
                            date_editor = WheelIgnoredDateEdit(self)
                            date_editor.setDate(date)
                            date_editor.setProperty('oldDate', date_editor.date())
                            date_editor.setCalendarPopup(True)
                            date_editor.dateChanged.connect(lambda new_date, editor=date_editor, row=row_index, col=col_index: self.create_date_changed_handler(editor, row, col)(new_date))
                            self.table.setCellWidget(row_index, col_index, date_editor)
                        if col_index == 12:
                            combo_box = self.table.cellWidget(row_index, 12)
                            combo_box.setCurrentText(row_data[12])
                        else:
                            item = NumericTableWidgetItem(value if value else '')
                            if col_index == 9 and not value:
                                item = NumericTableWidgetItem('Не применялась')
                                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                            if col_index == 10 and not value:
                                item = NumericTableWidgetItem('Отсутствуют')
                                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                            self.table.setItem(row_index, col_index, item)
                            self.table.resizeRowToContents(item.row())
                            item.setToolTip(value)
                        if col_index == 9 or col_index == 10:
                            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)

            for tab_name, table in self.other_tables.items():
                if tab_name == 'Информация':
                    self.load_and_display_excel_data()
                    self.load_information_data()
                else:
                    file_path = f'docs/{tab_name.lower().replace(" ", "_")}.csv'
                    with open(file_path, 'r', encoding='utf-8') as file:
                        table.setRowCount(0)
                        reader = csv.reader(file)
                        for row_data in reader:
                            if any(row_data):  # Only add rows that are not entirely empty
                                row_index = table.rowCount()
                                table.insertRow(row_index)
                                for col_index, value in enumerate(row_data):
                                    if "Unnamed:" in value:
                                        value = ""
                                    item = NumericTableWidgetItem(value if value else '')
                                    table.setItem(row_index, col_index, item)
                                    item.setToolTip(value)
                                    table.resizeRowToContents(item.row())
        except FileNotFoundError:
            QMessageBox.warning(self, 'Ошибка', 'Файл данных не найден.')
        finally:
            self.table.itemChanged.connect(self.item_changed)
            for _, tables in self.other_tables.items():
                tables.itemChanged.connect(self.item_changed)
            self.table.sortItems(0, QtCore.Qt.AscendingOrder)
            self.validate_all_dates()

    def load_information_data(self):
        """ Load the first two rows from информация.csv and set them as values for the 2nd and 3rd columns in the AOSR journal table. """
        try:
            with open('docs/информация.csv', 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                info_data = list(reader)[:2]  # Read only the first two rows

                for row_index in range(self.table.rowCount()):
                    # Ensure the table has enough rows to receive the data
                    if self.table.rowCount() < row_index + 1:
                        self.table.insertRow(self.table.rowCount())  # Add a new row if necessary

                    # Update the 2nd column with the first row of CSV data, if available
                    city_data = info_data[0][0]
                    self.table.setItem(row_index, 1, QTableWidgetItem(city_data))

                    # Update the 3rd column with the second row of CSV data, if available
                    object_name_data = info_data[1][0]
                    self.table.setItem(row_index, 2, QTableWidgetItem(object_name_data))

        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Не удалось загрузить информацию: {str(e)}')

    def reload_ov(self):
        with open('docs/вор.csv', 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            table = self.other_tables['ВОР']
            table.setRowCount(0)
            for row_index, row_data in enumerate(reader):
                if any(row_data):  # Check if the row is not entirely empty
                    table.insertRow(row_index)
                    for col_index, value in enumerate(row_data):
                        item = NumericTableWidgetItem(value if value else '')
                        table.setItem(row_index, col_index, item)
                        item.setToolTip(value)
                        table.resizeRowToContents(item.row())

    def reload_reg(self):
        with open('docs/реестр_ид.csv', 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            table = self.other_tables['Реестр ИД']
            table.setRowCount(0)
            for row_data in reader:
                if any(row_data):  # Ensure not to add empty rows
                    row_index = table.rowCount()
                    table.insertRow(row_index)
                    for col_index, value in enumerate(row_data):
                        item = NumericTableWidgetItem(value if value else '')
                        table.setItem(row_index, col_index, item)
                        item.setToolTip(value)
                        table.resizeRowToContents(item.row())

    def create_date_changed_handler(self, editor, row, col):
        def handle_date_changed(new_date):
            old_date = editor.property('oldDate') if editor.property('oldDate') is not None else editor.date()
            if old_date != new_date:
                command = EditDateCommand(editor, old_date, new_date)
                self.undo_stack.push(command)
            editor.setProperty('oldDate', new_date)  # Update the old date after pushing to undo stack
            self.save_changes()  # Save changes if necessary
        return handle_date_changed


    def resize_table_to_contents(self, table):
        table.resizeRowsToContents()
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)

    def load_and_display_excel_data(self):
        csv_path = 'docs/информация.csv'
        try:
            data_frame = pd.read_csv(csv_path, header=None)
            data_frame = data_frame.fillna('')  
            
            table = self.other_tables['Информация']
            table.setRowCount(len(data_frame.index))
            table.setColumnCount(len(data_frame.columns))
            table.horizontalHeader().setVisible(False)  
            editable_rows_first_column = {1, 2, 5, 8, 11, 15, 17}
            
            for index, row in data_frame.iterrows():
                for col_index, value in enumerate(row):
                    if "Unnamed:" in value:
                        value = ""
                    item = QTableWidgetItem(str(value))
                    item.setFlags(Qt.ItemIsEnabled)  
                    if col_index == len(row) - 1 or (col_index == 0 and (index + 1) in editable_rows_first_column):
                        item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled)
                        if index + 1 in [ 5,8,11,15 ] or col_index==len(row)-1:
                            item.setTextAlignment(Qt.AlignCenter)
                        item.setBackground(QColor(255, 255, 255))  
                    else:
                        item.setBackground(QColor(240, 240, 240))  

                    table.setItem(index, col_index, item)
            table.resizeColumnsToContents()
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Не удалось загрузить данные из Excel файла: {str(e)}')
    
    def date_item_changed(self, row, column, date):
        self.row_modified[row] = True
        self.validate_date(row, column)

    def item_changed(self, item):
        self.row_modified[item.row()] = True
        for _, items in self.other_tables.items():
            self.resize_table_to_contents(items)
        self.update_tooltip(item)

    def export_work_volume_to_general_ledger(self):
        with open('docs/вор.csv', 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            for row in range(self.table.rowCount()):
                if len(self.table.item(row, 7).text().split("_")) >= 4:
                    for item in self.table.item(row, 7).text().split("; "):
                        work_volume_data = item.split("_")
                        writer.writerow([work_volume_data[1], work_volume_data[3], work_volume_data[2], row+1])
                else:
                    writer.writerow("")
        self.reload_ov()

    def update_tooltip(self, item):
        item.setToolTip(item.text()) 

    def get_mtr_data(self):
    
        mtr_table = self.other_tables['Ведомость МТР']
        data = []
        for row in range(mtr_table.rowCount()):
            
            row_data = tuple(mtr_table.item(row, col).text() if mtr_table.item(row, col) else '' for col in range(mtr_table.columnCount()))
            data.append(row_data)
        return data

    def get_agreement_data(self):
    
        agreement_table = self.other_tables['Согласования']
        data = []
        for row in range(agreement_table.rowCount()):
            
            row_data = tuple(agreement_table.item(row, col).text() if agreement_table.item(row, col) else '' for col in range(agreement_table.columnCount()))
            data.append(row_data)
        return data
    
    def get_scheme_data(self):
    
        scheme_table = self.other_tables['Исп. схемы']
        data = []
        for row in range(scheme_table.rowCount()):
            
            row_data = tuple(scheme_table.item(row, col).text() if scheme_table.item(row, col) else '' for col in range(scheme_table.columnCount()))
            data.append(row_data)
        return data
    
    def add_record(self):
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)
        if row_count > 0:
            
            last_item = self.table.item(row_count - 1, 0)
            if last_item:
                last_aosr_number = int(last_item.text())
            else:
                last_aosr_number = 0
            next_aosr_number = last_aosr_number + 1
        else:
            next_aosr_number = 1  
        
        
        aosr_number_item = NumericTableWidgetItem(str(next_aosr_number))
        self.table.setItem(row_count, 0, aosr_number_item)
        self.set_default_value(row_count)
        self.table.setItem(row_count, 0, aosr_number_item)
        self.set_default_value(row_count)
        for col_index in range(1, self.table.columnCount()):
            if col_index in [3, 4, 5]:  
                date_editor = WheelIgnoredDateEdit(self)
                date_editor.setCalendarPopup(True)
                date_editor.setDate(QDate.currentDate())
                date_editor.dateChanged.connect(lambda new_date, editor=date_editor, row=row_count, col=col_index: self.create_date_changed_handler(editor, row, col)(new_date))
                date_editor.setProperty('oldDate', date_editor.date())
                self.table.setCellWidget(row_count, col_index, date_editor)
            else:
                item = NumericTableWidgetItem('')
                if col_index == 9:
                    item = NumericTableWidgetItem('Не применялась')
                    item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                if col_index == 10:
                    item = NumericTableWidgetItem('Отсутствуют')
                    item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                self.table.setItem(row_count, col_index, item)
        self.row_modified[row_count] = True

    def remove_record(self):
        current_row = self.table.currentRow()
        if current_row != -1:
            self.table.removeRow(current_row)
            self.save_all_tables()
        else:
            QMessageBox.warning(self, 'Ошибка', 'Выберите строку для удаления')

    def tab_changed(self, index):
        tab_name = self.tab_widget.tabText(index)
        self.setWindowTitle(f'АОСР - {tab_name}')

    def validate_date(self, row, column):
        date_editor = self.table.cellWidget(row, column)
        if not date_editor:
            return True  
        date = date_editor.date()
        if column == 3:  
            end_date_editor = self.table.cellWidget(row, 4)
            if end_date_editor and date > end_date_editor.date():
                date_editor.setStyleSheet("background-color: #ffcccc;")
                end_date_editor.setStyleSheet("background-color: #ffcccc;")
                return False
        elif column == 4:  
            start_date_editor = self.table.cellWidget(row, 3)
            if start_date_editor and date < start_date_editor.date():
                date_editor.setStyleSheet("background-color: #ffcccc;")
                start_date_editor.setStyleSheet("background-color: #ffcccc;")
                return False
        elif column == 5:  
            start_date_editor = self.table.cellWidget(row, 3)
            end_date_editor = self.table.cellWidget(row, 4)
            if (start_date_editor and date < start_date_editor.date()) or (end_date_editor and date < end_date_editor.date()):
                date_editor.setStyleSheet("background-color: #ffcccc;")
                return False
        date_editor.setStyleSheet("")  
        return True

    def validate_all_dates(self):
        for row in range(self.table.rowCount()):
            for col in [3, 4, 5]:  
                self.validate_date(row, col)

    def save_changes(self, filename='docs/журнал_аоср.csv'):
        with open(filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_ALL) 
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    if col in [3, 4, 5]:  
                        date_editor = self.table.cellWidget(row, col)
                        date = date_editor.date().toString("MM/dd/yyyy") if date_editor else ""
                        row_data.append(date)
                    elif col == 12:  
                        combo_box = self.table.cellWidget(row, col)
                        value = combo_box.currentText() if combo_box else 'Нет'
                        row_data.append(value)
                    else:
                        item = self.table.item(row, col)
                        text = item.text() if item else ""
                        row_data.append(text)
                writer.writerow(row_data)

    def save_table_data(self, table, filename):
        with open(filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_ALL) 
            for row in range(table.rowCount()):
                row_data = []
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    text = item.text() if item else ""
                    row_data.append(text)
                writer.writerow(row_data)

    def save_all_tables(self):
        try:
            for tab_name in self.tabs:
                if tab_name == 'Журнал АОСР':
                    self.save_changes()
                elif tab_name in self.other_tables:
                    table = self.other_tables[tab_name]
                    filename = f'{'docs/' + tab_name.lower().replace(" ", "_")}.csv'
                    self.save_table_data(table, filename)
            self.load_table_data()
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', 'Не удалось сохранить изменения.')        

    
    def open_volume_selection(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, 'Ошибка', 'Не выбрана строка для добавления объема работы.')
            return

        dialog = VolumeSelectionDialog(self)
        
        volume_data = self.get_volume_data()
        dialog.set_data(volume_data)
        
        if dialog.exec_() == QtWidgets.QDialog.Accepted and dialog.selected_data:
            
            self.reload_ov()
            self.table.setItem(selected_row, 7, QtWidgets.QTableWidgetItem(dialog.selected_data))
            self.save_changes()
            self.export_work_volume_to_general_ledger()  

    def get_volume_data(self):
        
        
        data = []
        volume_table = self.other_tables['Виды и объемы работ']
        for row in range(volume_table.rowCount()):
            row_data = []
            for col in range(volume_table.columnCount()):
                item = volume_table.item(row, col)
                row_data.append(item.text() if item else '')
            data.append(row_data)
        return data
    

def main():
    app = QApplication(sys.argv)
    ex = AOSRApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
