import sys
import csv
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import (QFileDialog, QComboBox, QHBoxLayout, QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox, QDateEdit, QHeaderView)
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QColor, QIcon

class SchemeSelectionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Выбрать исполнительную схему')
        self.layout = QtWidgets.QVBoxLayout(self)

        # Предполагаем, что в таблице три колонки
        self.table = QtWidgets.QTableWidget(0, 3)
        self.layout.addWidget(self.table)
        self.table.setHorizontalHeaderLabels(['Наименование схемы', '№', 'Примечание'])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)

        self.populate_table()

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

        self.table = QtWidgets.QTableWidget(0, 1)  # Предполагается, что таблица содержит 3 колонки
        self.layout.addWidget(self.table)
        self.table.setHorizontalHeaderLabels(['Наименование документа'])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)

        self.populate_table()

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

        # Setup the table to display MTR data with checkboxes
        self.table = QtWidgets.QTableWidget(0, 3)  # Assuming 4 columns for demonstration
        self.layout.addWidget(self.table)
        self.table.setHorizontalHeaderLabels(['Объект контроля', 'Сертификаты, паспорта и иные документы', 'Акты входного контроля'])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)  # Выбор по строкам
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)  # Разрешить выбор только одной строки за раз

        # Стилизация выделенных строк
        self.table.setStyleSheet("QTableWidget::item:selected { background-color: #add8e6; }")

        self.populate_table()

        # Кнопки OK и Cancel
        self.ok_button = QtWidgets.QPushButton('OK')
        self.ok_button.clicked.connect(self.accept_selection)
        self.cancel_button = QtWidgets.QPushButton('Отмена')
        self.cancel_button.clicked.connect(self.reject)
        
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        self.layout.addLayout(button_layout)
    

    def populate_table(self):
        mtr_data = self.parent().get_mtr_data()  # Получение данных из вкладки "Ведомость МТР"
        self.table.setRowCount(len(mtr_data))
        for row_index, row_data in enumerate(mtr_data):
            # Добавляем данные в каждую колонку таблицы
            for col_index in range(0, self.table.columnCount()):
                self.table.setItem(row_index, col_index, QtWidgets.QTableWidgetItem(row_data[col_index+1]))


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


        # Append the selection to the main journal table
        selected_row = self.parent().table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, 'Ошибка', 'Не выбрана строка в журнале.')
            return

        existing_item = self.parent().table.item(selected_row, 9)
        existing_text = existing_item.text() if existing_item and existing_item.text() != "Не применялась" else ""
        new_text = "; ".join(selected_items)
        final_text = existing_text + "; " + new_text if existing_text else new_text
        self.parent().table.setItem(selected_row, 9, QtWidgets.QTableWidgetItem(final_text))
        self.accept()


class VolumeSelectionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Выбрать объём работы')
        self.layout = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableWidget(0, 3)  # Таблица для данных "Виды и объемы работ"
        self.layout.addWidget(self.table)
        
        self.ok_button = QtWidgets.QPushButton('OK')
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button = QtWidgets.QPushButton('Отмена')
        self.cancel_button.clicked.connect(self.reject)
        
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        self.layout.addLayout(button_layout)

        self.selected_data = None

        self.table.cellDoubleClicked.connect(self.cell_double_clicked)

    def cell_double_clicked(self, row, column):
        # Prompt the user to input a number
        num, ok = QtWidgets.QInputDialog.getInt(self, "Введите объем", "Объем работы:")
        if ok:
            # Fetch existing data in the 7th column of the main table
            selected_row = self.parent().table.currentRow()
            existing_data_item = self.parent().table.item(selected_row, 6)

            # If there's already data in the cell, get it, otherwise, start with an empty string
            existing_data = existing_data_item.text() if existing_data_item else ""
            
            # Format the new data with the inputted number
            work_type = self.table.item(row, 0).text()
            volume = self.table.item(row, 1).text()
            unit = self.table.item(row, 2).text() if self.table.item(row, 2) else ""
            
            # Create the new selection string
            new_selection = f"{unit}_{work_type}_{num}_{volume}"

            # Append the new data to the existing data, ensuring it doesn't just duplicate the existing content
            if existing_data:
                self.selected_data = f"{existing_data}; {new_selection}".strip()
            else:
                self.selected_data = new_selection
        
            try:
                with open('docs/общая_ведомость.csv', 'a+', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    file.seek(0)
                    reader = csv.reader(file)

                    _, work_type, num, volume = new_selection.split("_")
                    aosr_number = self.parent().table.item(self.parent().table.currentRow(), 0).text()
                    
                    writer.writerow([work_type, volume, num, aosr_number])
            except Exception as e:
                QMessageBox.warning(self, 'Ошибка', f'Не удалось сохранить данные: {e}')
            self.accept()  # Close the dialog with confirmation of the selection
        else:
            self.selected_data = None
            self.reject()  # Close the dialog without making a selection

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

class AOSRApp(QMainWindow):
    def __init__(self):
        super().__init__()
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
            ('Общая ведомость', 'images/1.png'),
            ('Исп. схемы', 'images/3.png'),
            ('Согласования', 'images/6.png'),
            ('Виды и объемы работ', 'images/5.png'),
            ('Реестр ИД', 'images/4.png'),
        ]
        for name, icon_path in tab_names:
            tab = QWidget()
            self.tabs[name] = tab
            self.tab_widget.addTab(tab, QIcon(icon_path), name)  # Установка иконки и пустого текста

        self.setupJournalAOSRTab()
        self.setupOtherTabs()
        self.load_table_data()

        self.table.itemChanged.connect(self.item_changed)

    def setupJournalAOSRTab(self):
        layout = QVBoxLayout()
        self.tabs['Журнал АОСР'].setLayout(layout)
        self.table = QTableWidget(0, 14)  
        # Updating headers to include the new 'Наименование объекта' column and adjust others
        self.table.setHorizontalHeaderLabels([
            'Номер\nАОСР', 'Город\nстроительства', 'Наименование объекта', 'Дата начала\nработ', 'Дата окончания\nработ',
            'Дата подписания\nакта', '1. К освидетельствованию\nпредъявлены работы', 'Вид и объём работ\n(выгрузка ведомости)',
            'Работы выполнены\nпо проектно-сметной \nдок.', 'При\nвыполнении работ\nприменены', 'При\nвыполнении работ\nотклонения',
            'Основание для\nпроизводства работ', 'Сформировать\nакт', 'Исполнительная\nсхема'
        ])
        self.table.resizeRowsToContents()  # Resizes rows based on content automatically
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setStretchLastSection(True)
        self.initialize_default_values()
        for column in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(column)
            header_item.setToolTip(header_item.text().replace('\n', ' '))  # Set tooltip to be the header text without new lines

        # Create a button panel
        button_panel = QWidget()
        button_layout = QHBoxLayout()
        button_panel.setLayout(button_layout)
        
        # Add buttons
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

        button_layout.addWidget(btn_create_act)
        button_layout.addWidget(btn_select_mtr)
        button_layout.addWidget(btn_select_volume)
        button_layout.addWidget(btn_agreements)
        button_layout.addWidget(btn_select_scheme)
        layout.addWidget(button_panel)
        layout.addWidget(self.table)

        self.table.itemChanged.connect(self.item_changed)
        btn_add = QPushButton('Добавить запись')
        btn_remove = QPushButton('Удалить запись')
        btn_add.clicked.connect(self.add_record)
        btn_remove.clicked.connect(self.remove_record)
        layout.addWidget(btn_add)
        layout.addWidget(btn_remove)
    def open_mtr_selection(self):
        dialog = MTRSelectionDialog(self)
        dialog.exec_()
    def setupOtherTabs(self):
        self.other_tables = {}
        tab_configs = {
            'Исп. схемы': ['Наименование схем', '№', ''],
            'Согласования': ['Наименование документа'],
            'Виды и объемы работ': ['', 'Ед.изм', 'Номер'],
            'Реестр ИД': ['№п/п', '№ акта', 'Наименование', 'Кол-во листов', 'Примечание'],
            'Ведомость МТР': ['п/п', 'Объект контроля', 'Сертификаты, паспорта и иные документы', '*Акты входного контроля'],
            'Общая ведомость': ['Наименование выполненных работ', 'Ед.изм', 'Кол-во', 'Примечание']
        }

        for name, headers in tab_configs.items():
            tab = self.tabs[name]
            layout = QVBoxLayout()
            table = QTableWidget(0, len(headers))
            table.resizeRowsToContents()  # Resizes rows based on content automatically
            table.setHorizontalHeaderLabels(headers)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.horizontalHeader().setStretchLastSection(True)
            self.initialize_default_values()
            for column in range(table.columnCount()):
                header_item = table.horizontalHeaderItem(column)
                header_item.setToolTip(header_item.text())  # Tooltips for headers

            tab.setLayout(layout)
            layout.addWidget(table)

            button_panel = QWidget()
            button_layout = QHBoxLayout()
            button_panel.setLayout(button_layout)

            btn_add = QPushButton('Добавить запись')
            btn_remove = QPushButton('Удалить запись')
            btn_save = QPushButton('Сохранить записи')
            btn_add.clicked.connect(lambda checked, t=table: self.add_other_record(t))
            btn_remove.clicked.connect(lambda checked, t=table: self.remove_other_record(t))
            btn_save.clicked.connect(lambda: self.save_all_tables())

            button_layout.addWidget(btn_add)
            button_layout.addWidget(btn_remove)
            button_layout.addWidget(btn_save)
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

    def set_default_value(self, row_index):
    # Устанавливаем значения по умолчанию для столбцов 8 и 9
        item_8 = QTableWidgetItem('Не применялась')
        item_9 = QTableWidgetItem('Отсутствуют')
        combo_box_11 = QtWidgets.QComboBox()
        combo_box_11.addItems(['Да', 'Нет'])  # Добавляем варианты выбора
        combo_box_11.setCurrentText('Да')

        # Устанавливаем флаги, чтобы сделать ячейки нередактируемыми
        item_8.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        item_9.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)

        # Добавляем элементы в таблицу
        self.table.setItem(row_index, 9, item_8)
        self.table.setItem(row_index, 10, item_9)
        self.table.setCellWidget(row_index, 12, combo_box_11)

    def export_to_pdf_and_xls(self):
        pass

    def tab_changed(self, index):
        tab_name = self.tab_widget.tabText(index)
        self.setWindowTitle(f'АОСР - {tab_name}')

    def load_table_data(self):
        self.table.itemChanged.disconnect(self.item_changed)
        try:
            with open('docs/журнал_аоср.csv', 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                for row_index, row_data in enumerate(reader):
                    self.table.insertRow(row_index)
                    self.set_default_value(row_index)
                    for col_index, value in enumerate(row_data):
                        if col_index in [3, 4, 5]:  # Date columns
                            date = QDate.fromString(value, "yyyy-MM-dd") if value else QDate.currentDate()
                            date_editor = QDateEdit(self)
                            date_editor.setDate(date)
                            date_editor.setCalendarPopup(True)
                            date_editor.dateChanged.connect(self.create_date_changed_handler(row_index, col_index))
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
                if tab_name in ['Журнал АОСР', 'Информация']:  # Skip these tabs
                    continue
                with open(f'{'docs/' + tab_name.lower().replace(" ", "_")}.csv', 'r', encoding='utf-8') as file:
                    reader = csv.reader(file)
                    for row_index, row_data in enumerate(reader):
                        table.insertRow(row_index)
                        for col_index, value in enumerate(row_data):
                            item = NumericTableWidgetItem(value if value else '')
                            table.setItem(row_index, col_index, item)
                            item.setToolTip(value)
                            table.resizeRowToContents(item.row())
        except FileNotFoundError:
            QMessageBox.warning(self, 'Ошибка', 'Файл данных не найден.')
        finally:
            self.table.itemChanged.connect(self.item_changed)
            self.table.sortItems(0, QtCore.Qt.AscendingOrder)  # Sort by AOSR number after loading
            self.validate_all_dates()

    def reload_ov(self):
        with open('docs/общая_ведомость.csv', 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                table = self.other_tables['Общая ведомость']
                table.setRowCount(0)
                for row_index, row_data in enumerate(reader):
                    table.insertRow(row_index)
                    for col_index, value in enumerate(row_data):
                        item = NumericTableWidgetItem(value if value else '')
                        table.setItem(row_index, col_index, item)
                        item.setToolTip(value)
                        table.resizeRowToContents(item.row())
    def create_date_changed_handler(self, row, col):
        def handle_date_changed(date):
            self.date_item_changed(row, col, date)
            self.save_all_tables()
        return handle_date_changed

    def date_item_changed(self, row, column, date):
        self.row_modified[row] = True
        self.validate_date(row, column)

    def item_changed(self, item):
        self.row_modified[item.row()] = True
        self.table.resizeRowToContents(item.row())
        self.update_tooltip(item)

    def update_tooltip(self, item):
        item.setToolTip(item.text()) 

    def get_mtr_data(self):
    # Этот метод извлекает данные из таблицы "Ведомость МТР" и возвращает их в виде списка кортежей
        mtr_table = self.other_tables['Ведомость МТР']
        data = []
        for row in range(mtr_table.rowCount()):
            # Извлекаем данные из каждой строки: п/п, Объект контроля, Сертификаты, паспорта и иные документы, *Акты входного контроля
            row_data = tuple(mtr_table.item(row, col).text() if mtr_table.item(row, col) else '' for col in range(mtr_table.columnCount()))
            data.append(row_data)
        return data

    def get_agreement_data(self):
    # Этот метод извлекает данные из таблицы "Ведомость МТР" и возвращает их в виде списка кортежей
        agreement_table = self.other_tables['Согласования']
        data = []
        for row in range(agreement_table.rowCount()):
            # Извлекаем данные из каждой строки: п/п, Объект контроля, Сертификаты, паспорта и иные документы, *Акты входного контроля
            row_data = tuple(agreement_table.item(row, col).text() if agreement_table.item(row, col) else '' for col in range(agreement_table.columnCount()))
            data.append(row_data)
        return data
    
    def get_scheme_data(self):
    # Этот метод извлекает данные из таблицы "Ведомость МТР" и возвращает их в виде списка кортежей
        scheme_table = self.other_tables['Исп. схемы']
        data = []
        for row in range(scheme_table.rowCount()):
            # Извлекаем данные из каждой строки: п/п, Объект контроля, Сертификаты, паспорта и иные документы, *Акты входного контроля
            row_data = tuple(scheme_table.item(row, col).text() if scheme_table.item(row, col) else '' for col in range(scheme_table.columnCount()))
            data.append(row_data)
        return data
    
    def add_record(self):
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)
        if row_count > 0:
            # Get the last row's Номер АОСР and increment it
            last_item = self.table.item(row_count - 1, 0)
            if last_item:
                last_aosr_number = int(last_item.text())
            else:
                last_aosr_number = 0
            next_aosr_number = last_aosr_number + 1
        else:
            next_aosr_number = 1  # Start from 1 if the table is empty
        
        # Setting the next Номер АОСР
        aosr_number_item = NumericTableWidgetItem(str(next_aosr_number))
        self.table.setItem(row_count, 0, aosr_number_item)
        self.set_default_value(row_count)
        self.table.setItem(row_count, 0, aosr_number_item)
        self.set_default_value(row_count)
        for col_index in range(1, self.table.columnCount()):
            if col_index in [3, 4, 5]:  # For date columns
                date_editor = QDateEdit(self)
                date_editor.setCalendarPopup(True)
                date_editor.setDate(QDate.currentDate())
                date_editor.dateChanged.connect(self.create_date_changed_handler(row_count, col_index))
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
            return True  # No date editor, no validation needed
        date = date_editor.date()
        if column == 3:  # Start date
            end_date_editor = self.table.cellWidget(row, 4)
            if end_date_editor and date > end_date_editor.date():
                date_editor.setStyleSheet("background-color: #ffcccc;")
                end_date_editor.setStyleSheet("background-color: #ffcccc;")
                return False
        elif column == 4:  # End date
            start_date_editor = self.table.cellWidget(row, 3)
            if start_date_editor and date < start_date_editor.date():
                date_editor.setStyleSheet("background-color: #ffcccc;")
                start_date_editor.setStyleSheet("background-color: #ffcccc;")
                return False
        elif column == 5:  # Sign date
            start_date_editor = self.table.cellWidget(row, 3)
            end_date_editor = self.table.cellWidget(row, 4)
            if (start_date_editor and date < start_date_editor.date()) or (end_date_editor and date < end_date_editor.date()):
                date_editor.setStyleSheet("background-color: #ffcccc;")
                return False
        date_editor.setStyleSheet("")  # Reset style if no errors
        return True

    def validate_all_dates(self):
        for row in range(self.table.rowCount()):
            for col in [3, 4, 5]:  # Validate all date columns
                self.validate_date(row, col)

    def save_changes(self, filename='docs/журнал_аоср.csv'):
        with open(filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    if col in [3, 4, 5]:  # Date columns
                        date_editor = self.table.cellWidget(row, col)
                        date = date_editor.date().toString("yyyy-MM-dd") if date_editor else ""
                        row_data.append(date)
                    elif col == 12:  # Для столбца актов
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
            writer = csv.writer(file)
            for row in range(table.rowCount()):
                row_data = []
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    text = item.text() if item else ""
                    row_data.append(text)
                writer.writerow(row_data)

    def save_all_tables(self):
    # Сохраняем данные из основной таблицы Журнал АОСР
        for tab_name in self.tabs:
            if tab_name == 'Журнал АОСР':
                self.save_changes()
            elif tab_name in self.other_tables:
                table = self.other_tables[tab_name]
                filename = f'{'docs/' + tab_name.lower().replace(" ", "_")}.csv'
                self.save_table_data(table, filename)

    
    def open_volume_selection(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, 'Ошибка', 'Не выбрана строка для добавления объема работы.')
            return

        dialog = VolumeSelectionDialog(self)
        # Передача данных в диалоговое окно
        volume_data = self.get_volume_data()
        dialog.set_data(volume_data)
        
        if dialog.exec_() == QtWidgets.QDialog.Accepted and dialog.selected_data:
            # Вставляем данные в 7-й столбец выбранной строки
            self.reload_ov()
            self.table.setItem(selected_row, 7, QtWidgets.QTableWidgetItem(dialog.selected_data))

    def get_volume_data(self):
        # Этот метод должен возвращать данные из вкладки "Виды и объемы работ" в виде списка списков
        # Пример: [['Разработка грунта в котловане', '213', 'м3'], ...]
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
