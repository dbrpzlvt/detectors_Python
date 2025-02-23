'''
Подготовка данных от детекторов к последующей обработке
'''
import glob

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

import time
import numpy as np
import pandas as pd
import openpyxl
import re
from re import search
import os
import sys
import xlwings as xw
from xlwings.utils import rgb_to_int, int_to_rgb
import logging
import pyscreenshot
from tqdm import tqdm
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox

import check_correct


class Application:

    def __init__(self, parent):
        self.bs_intensivnosti_full = pd.DataFrame()
        self.bs_SSID_full = pd.DataFrame()
        self.bs_intensivnosti = []
        self.bs_SSID = []
        self.parent = parent
        self.parent.title("Добро пожаловать в приложение!")
        self.parent.geometry("1200x680")

        self.main_frame = tk.Frame(self.parent)
        self.main_frame.grid(sticky="nsew")
        self.create_widgets()

        self.cur_year = '2024'
        self.directory_raw = f'../raw_data/{self.cur_year}/Исходные данные'
        self.directory_pre = f'../raw_data/{self.cur_year}/Первичная обработка'
        self.directory_imp = f'../raw_data/{self.cur_year}/Импортированные данные'
        self.directory_pic = f'../raw_data/{self.cur_year}/Графики'
        self.folder_path = None

        # вызвал экземпляр класса Checking с проверкой файлов из скрипта check_correct.py
        self.checking = check_correct.Checking(self.editor, self.parent)

        # логи
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)

        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s %(message)s',
                            datefmt='%a, %d %b %Y %H:%M:%S',
                            filename='log.txt',
                            filemode='a+')

    def create_widgets(self):
        # Create a Label to display an instruction
        self.label = Label(self.parent, text="Обрабатывает файлы замеров, поступивших от ...\n"
                                             "Выберите директорию с исходными файлами или папками с файлами\n"
                                             "Для обратки только что поступивших файлов датчиков можно выбрать папку 'Исходные данные', даже если там находятся подпапки и далее выделить нужные файлы/папки и нажать запустить\n",
                                             # "Для обработки файлов, с проверенной структурой можно выбрать папку 'Первичная обработка' (с префиксом PRE_)\n"
                                             # "Для отрисовки графиков выбрать папку 'Импортированные данные' (с префиксом IN_)",
                           anchor="w", justify=LEFT)
        self.label.grid(row=0, padx=10, pady=10, columnspan=7, sticky=E)

        # Create a Listbox to display folder contents
        self.folder_contents = tk.Listbox(self.parent, selectmode=tk.EXTENDED, exportselection=False)
        self.folder_contents.grid(row=2, column=0, columnspan=5, sticky=NSEW, padx=10)  # .pack(fill=tk.BOTH, expand=True)

        # Create a scrollbar for the Listbox
        self.scrollbar = tk.Scrollbar(self.folder_contents)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.folder_contents.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.configure(command=self.folder_contents.yview)

        # Create a button to browse for a folder
        self.browse_button = tk.Button(self.parent, text="Browse Folder", command=self.browse_folder)
        self.browse_button.grid(row=1, column=0, sticky=W, pady=10, padx=10)  # .pack(pady=10)

        # Create a Button to select all items in ListBox container
        self.select_button = tk.Button(self.parent, text="Select All", command=self.select_all)
        self.select_button.grid(row=3, column=2, sticky=W, pady=10, padx=10)  # .pack(pady=10)
        self.editor = ScrolledText(self.parent, wrap='word')
        self.editor.grid(row=2, column=6, padx=(10, 10), sticky=NE)

        # Create a ComboBox to show a menu
        self.combobox = Combobox(self.parent, values=['Проверка структуры файлов',
                                                      'Предобработка',
                                                      'Вставка обработанных данных в шаблон .xlsm',
                                                      'Скриншоты'],
                                 width=30, state="readonly")
        self.combobox.grid(row=3, column=0, pady=10, padx=10)
        self.combobox.bind("<<ComboboxSelected>>", self.DeleteOptions)

        # Create a Label to show info for RadioButton
        self.lbl = Label(self.parent, text="Проверять тип файла перед обработкой?")

        # Create a RadioButton to ability making choice
        self.var = tk.StringVar()
        self.var.set("yes")
        self.yes_rdb = Radiobutton(self.parent, text='Да', variable=self.var, value="yes")
        self.no_rdb = Radiobutton(self.parent, text='Нет', variable=self.var, value="no")
        self.yes_rdb.grid_forget()
        self.no_rdb.grid_forget()

        # Главная кнопочка
        self.run_btn = Button(
            self.parent,
            text="Запустить",
            command=lambda: self.selected(self.combobox.get(), self.var.get())
        )
        self.run_btn.grid(row=5, column=2, padx=10, pady=10)
        # self.make_dynamic(self.main_frame)

    # def make_dynamic(self, widget):
    #     col_count, row_count = widget.grid_size()
    #
    #     for i in range(row_count):
    #         widget.grid_rowconfigure(i, weight=1)
    #
    #     for i in range(col_count):
    #         widget.grid_columnconfigure(i, weight=1)
    #
    #     for child in widget.children.values():
    #         child.grid_configure(sticky="nsew")
    #         self.make_dynamic(child)

    # Метод для добавления списка содержимого папки в ListBox
    def browse_folder(self):
        self.folder_path = filedialog.askdirectory()  # Open a folder selection dialog
        print(self.folder_path)
        if self.folder_path:
            self.folder_contents.delete(0, tk.END)  # Clear the Listbox
            for item in os.listdir(self.folder_path):
                self.folder_contents.insert(tk.END, item)  # Insert folder contents into Listbox

    # Метод для выбора всего содержимого в ListBox
    def select_all(self):
        self.folder_contents.select_set(0, tk.END)

    def DeleteOptions(self, _):
        selected_value = self.combobox.get()
        if selected_value == "Проверка структуры файлов" or selected_value == 'Скриншоты':
            self.lbl.grid_forget()
            self.yes_rdb.grid_forget()
            self.no_rdb.grid_forget()
            # self.hideDeleteOptions()
        elif selected_value == 'Предобработка' or selected_value == 'Вставка обработанных данных в шаблон .xlsm':

            self.lbl.grid(row=4, column=0, padx=10)  # .pack(pady=10)
            self.yes_rdb.grid(row=5, column=0, padx=10, sticky=W)
            self.no_rdb.grid(row=6, column=0, padx=10, sticky=W)
            self.var.set("yes")
            print('Установлено - yes')
            self.parent.update_idletasks()
            # time.sleep(2)
            # self.parent.update_idletasks()
            # self.showDeleteOptions()
            # self.var.set("yes")

    # def hideDeleteOptions(self):
    #     self.lbl.grid_forget()
    #     self.yes_rdb.grid_forget()
    #     self.no_rdb.grid_forget()

    # def showDeleteOptions(self):
    #     self.lbl.grid(row=4, column=0, padx=10)  # .pack(pady=10)
    #     self.yes_rdb.grid(row=5, column=0, padx=10, sticky=W)
    #     self.no_rdb.grid(row=6, column=0, padx=10, sticky=W)

    def selected(self, event, rbt_var):

        selection = self.combobox.get()
        tmp = [self.folder_contents.get(idx) for idx in
               self.folder_contents.curselection()]  # список папок или файлов из Combobox в окошке
        if not tmp:
            tk.messagebox.showerror(title="ALERT",
                                    message="Не выбран ни один файл или папка из списка.\nВыберите и повторите попытку")
            return
        # print(tmp)
        xlsx_files = []  # список файлов, если в папке "Исходные данные" не файлы, а аще папки
        for i in tmp:  # для каждого файла или папки из Combobox проверяем файл это или папка
            if os.path.isfile(os.path.join(self.folder_path, i)):
                xlsx_files.append(i)  # если файл, то все ок - оставляем список как есть
                # folder_before = os.path.basename(os.path.dirname(os.path.join(self.folder_path, i)))
                # print(folder_before)
                # if folder_before != 'Исходные данные':
                #     xlsx_files.append('/' + folder_before + '/' + i) # если файл внутри подпапки (обычно это название дороги), то добавляем к имени файла еще папку в которой он лежет
                #     print(xlsx_files)
                # elif folder_before in ['Первичная обработка', 'Импортированные данные']:
                #     xlsx_files.append(i)  # если файл, то все ок - оставляем список как есть
            else:
                for file in os.listdir(os.path.join(self.folder_path, i)):
                    xlsx_files.append(os.path.join(i,
                                                   file))  # если папка, не ок - открываем каждую папку и смотрим что там за файл и добавляем в новый список

        if selection == "Проверка структуры файлов":
            self.editor.delete(tk.END)
            print(f'Выбор 1.\nПроверяю файлы {chr(10).join(xlsx_files)} на соответствие структуре\n')
            self.editor.insert(tk.END, f'\nВыбор 1.\nПроверяю файлы {chr(10).join(xlsx_files)} на соответствие структуре\n')
            for filename in tqdm(xlsx_files):
                self.editor.insert(tk.END, '\n======' + filename + '======\n')
                f = os.path.join(self.folder_path, filename)
                if os.path.isfile(f):
                    self.parent.update_idletasks()
                    time.sleep(2)
                    self.structure_check(f)
        elif selection == "Предобработка":
            # self.var.set("yes")
            print(f'Выбор 2.\nПредобрабатываю файлы {chr(10).join(xlsx_files)}...\n')
            self.editor.delete(tk.END)
            self.parent.update_idletasks()
            self.editor.insert(tk.END, f'\nВыбор 2.\nПредобрабатываю файлы {chr(10).join(xlsx_files)}...\n')
            for filename in tqdm(xlsx_files):
                self.editor.see(END)
                self.editor.insert(tk.END, f'\n======' + filename + '======\n')
                f = os.path.join(self.folder_path, filename)
                if os.path.isfile(f):
                    self.parent.update_idletasks()
                    time.sleep(2)
                    if rbt_var == 'yes':
                        self.editor.insert(tk.END, f'Включена проверка файла на соотвествтие структуре...\n'
                                                   f'Проверяю сначала структуру файла {filename}\n')
                        self.preprocessing(filename, self.var)
                    elif rbt_var == 'no':
                        self.preprocessing(filename)
                self.editor.insert(tk.END, f'\n{filename} обработан!\n')
        elif selection == "Вставка обработанных данных в шаблон .xlsm":
            # self.var.set("yes")
            # self.combobox.bind("<<ComboboxSelected>>", lambda: self.showDeleteOptions())
            print(f'\nВыбор 3.\nВставляю данные файлов {chr(10).join(xlsx_files)} в шаблон отчета...\n')
            self.editor.delete(tk.END)
            self.editor.insert(tk.END, f'\nВыбор 3.\nВставляю данные файлов {chr(10).join(xlsx_files)} в шаблон отчета...\n')
            for filename in tqdm(xlsx_files):
                self.editor.insert(tk.END, '\n======' + filename + '======\n')
                self.editor.see(END)
                f = os.path.join(self.folder_path, filename)
                # checking if it is a file
                if os.path.isfile(f):
                    self.parent.update_idletasks()
                    time.sleep(2)
                    if rbt_var == 'yes':
                        self.editor.insert(tk.END, f'Включена проверка файла на соотвествтие структуре...\n'
                                                   f'Проверяю сначала структуру файла {filename}\n')
                        self.init_data_import(filename, self.var)
                    elif rbt_var == 'no':
                        self.init_data_import(filename)
                self.editor.insert(tk.END, f'Файл {filename} обработан!\n')
                self.editor.see(END)
        elif selection == "Скриншоты":
            print(f'\nВыбор 4.\nСоздаю скриншоты файлов {chr(10).join(xlsx_files)}...\n')
            self.editor.delete(tk.END)
            self.editor.insert(tk.END, f'\nВыбор 4.\nСоздаю скриншоты файлов {chr(10).join(xlsx_files)}...\n')
            for filename in tqdm(xlsx_files):
                self.editor.insert(tk.END, '\n======' + filename + '======\n')
                self.editor.see(END)
                f = os.path.join(self.folder_path, filename)
                # checking if it is a file
                if os.path.isfile(f):
                    self.parent.update_idletasks()
                    time.sleep(2)
                    self.screenshots(filename)
                self.editor.insert(tk.END, f'{filename} обработан!\n')

    def structure_check(self, file):
        # print(f'\n\n====== FILENAME FROM FUNCTION {os.path.join(directory_raw, file)}\n\n')
        # file = 'PRE_км 41+138 а.д М-10 Россия Москва – Санкт-Петербург .xlsx'
        filename_prefix = re.match('^PRE_.+', str(file))
        if filename_prefix:
            wb = openpyxl.load_workbook(os.path.join(self.directory_pre, file))
        else:
            wb = openpyxl.load_workbook(os.path.join(self.folder_path, file))
        sheet = wb.active

        # 1. Проверка соответствия отчёта заданной форме (с записью итогов проверки в лог-файл)

        if sheet['B3'].value == "Общая интенсивность автомобилей" and sheet['E3'].value == "Легковые (до 6 м)" and \
                sheet[
                    'H3'].value == "Малые груз. (6-9 м)" and sheet['K3'].value == "Грузовые (9-13 м)" and \
                sheet['N3'].value == "Груз. большие (13-22 м)" and sheet['Q3'].value == "Автопоезда (22-30 м)" and \
                sheet[
                    'T3'].value == "Автобусы" and sheet['W3'].value == "Мотоциклы":
            if filename_prefix:
                logging.info(f'{file}: Формат отчёта соответствует 7 категориям ТС, Варианту №1 - (шаблон sample_fda_var1.xlsm)')
                self.editor.insert(tk.END,
                              f'{file}: Формат отчёта соответствует 7 категориям ТС, Варианту №1\nИспользую шаблон отчета (sample_fda_var1.xlsm)\n')

            else:
                logging.info(f'{file}: Формат отчёта соответствует 7 категориям ТС, Варианту №1 - (шаблон pre_sample_r1.xlsx)')
                self.editor.insert(tk.END,
                              f'{file}: Формат отчёта соответствует 7 категориям ТС, Варианту №1\nИспользую шаблон отчета (pre_sample_r1.xlsx)\n')
            return 'Rosautodor_1'
        elif sheet['B3'].value == "Общая интенсивность автомобилей" and sheet['E3'].value == "Легковые (до 4.5 м)" and \
                sheet[
                    'H3'].value == "Легковые большие (4-6 м)" and sheet[
            'K3'].value == "Малые груз. (6-9 м)" and sheet['N3'].value == "Грузовые (9-13 м)" and \
                sheet['Q3'].value == "Груз. большие (13-22 м)" and sheet['T3'].value == "Автопоезда (22-30 м)" and \
                sheet[
                    'W3'].value == "Автобусы" and sheet['Z3'].value == "Мотоциклы":
            if filename_prefix:
                logging.info(f'{file}: Формат отчёта соответствует 7 категориям ТС, Варианту №2 - (шаблон sample_fda_var2.xlsm)')
                self.editor.insert(tk.END,
                              f'{file}: Формат отчёта соответствует 7 категориям ТС, Варианту №2\nИспользую шаблон отчета (sample_fda_var2.xlsm)\n')
            else:
                logging.info(f'{file}: Формат отчёта соответствует 7 категориям ТС, Варианту №2 - (шаблон pre_sample_r2.xlsx)')
                self.editor.insert(tk.END,
                              f'{file}: Формат отчёта соответствует 7 категориям ТС, Варианту №2\nИспользую шаблон отчета (pre_sample_r2.xlsx)\n')

            return 'Rosautodor_2'

        # elif sheet['B6'].value=="За период" and sheet['E7'].value=="легковые автомобили (до 6 м)" and sheet['H7'].value=="микроавтобусы, малые грузовые автомобили (6-9 м)" and sheet['K7'].value=="грузовые автомобили (9-11 м)" and\
        #         sheet['N7'].value=="автобусы (11-13 м)" and sheet['Q7'].value=="грузовые большие автомобили, автопоезда (13-18 м)" and sheet['T7'].value=="длинные автопоезда (> 18 м)":
        #     logging.info(f'{filename}: Формат отчёта соответствует 6 категориям ТС, Варианту №2')
        # elif sheet['B6'].value == "За период" and sheet['E7'].value == "легковые автомобили (до 6 м)" and sheet['H7'].value == "малые грузовые автомобили до 5 тонн (6-9 м)" and\
        #         sheet['K7'].value == "грузовые автомобили 5-12 тонн (9-11 м)" and sheet['N7'].value == "автобусы (11-13 м)" and sheet['Q7'].value == "грузовые большие автомобили 12-20 тонн (13-22 м)" and\
        #         sheet['T7'].value == "автопоезда более 20 тонн (22-30 м)":
        #     logging.info(f'{filename}: Формат отчёта соответствует 6 категориям ТС, Варианту №3')
        #     editor.insert(tk.END, f'{file}: Формат отчёта соответствует 6 категориям ТС, Варианту №3')
        else:
            logging.info(f'{file}: Формат отчёта не совпадает')
            self.editor.insert(tk.END, f'{file}: Формат отчёта не совпадает')
            return 0

    def preprocessing(self, file, rbt_var=None):

        wb = openpyxl.load_workbook(os.path.join(self.folder_path, file))  # открываем файл с данными
        sheet = wb.active  # активируем лист
        columnA = sheet['A']  # запоминаем колонку А

        dataset_begins = []  # список ячеек, где хранятся текстовые данные для аккуратных наименований

        for cell in columnA:
            if (search("км", str(cell.value)) or search("km", str(cell.value))) and cell.row > 4:
                dataset_begins = dataset_begins + [cell.row]
        dataset_begins = dataset_begins + [len(columnA) + 1]

        type_file = None
        # Старый коммента - # Структуру исходных файлов не проверяю, поскольку сделал это уже ранее. Единый формат отчёта
        # Теперь структуру проверяем (можно выбрать проверять или нет с помощью RadioButton)
        if rbt_var:  # если RadioButton выбран "да", то происходит проверка структуры и в type_file возвращается тип структуры 'Rosautodor_1' или 'Rosautodor_2'
            type_file = self.structure_check(file)

        wb_data = xw.Book(os.path.join(self.folder_path, file))
        wb_sample = None
        # Копирую данные о названии дороги и местоположении детектора, категориях ТС
        for i in range(0, len(dataset_begins) - 1):
            # ниже старые комменты
            # Проверка соответствия структуры исходных файлов отключена, поскольку сделал такую проверку ранее отдельно.
            # В текущей версии все файлы от Росавтодора - единого формата.

            # теперь если включена проверка структуры, то в зависимости от того, что вернулось, тот шаблон и открываем и вставляем туда данные
            if type_file:
                if type_file == 'Rosautodor_1':
                    wb_sample = xw.Book('../pre_sample_r1.xlsx')  # pre_sample_r1.xlsx
                elif type_file == 'Rosautodor_2':
                    wb_sample = xw.Book('../pre_sample_r2.xlsx')  # pre_sample_r2.xlsx
            else:
                tk.messagebox.showwarning(title="ALERT",
                                          message="Вы выбрали не проверять структуру файла. Выберите файл шаблона вручную\n"
                                                  "Выберите файл шаблона (pre_sample_r1.xlsx) или (pre_sample_r2.xlsx)\n"
                                                  "Посмотреть какой именно шаблон подходит к обрабатываемому файлу можно посмотреть в log-файле в корне проекта")
                folder_path_to_sample = filedialog.askopenfile()  # Open a folder selection dialog
                print(folder_path_to_sample)
                if folder_path_to_sample:
                    wb_sample = xw.Book(folder_path_to_sample.name)  # pre_sample_r2.xlsx
                elif not folder_path_to_sample:
                    tk.messagebox.showerror(title="ALERT",
                                            message="Не выбран ни один файл шаблона.\nВыберите и повторите попытку")
                    print("Не выбран ни один файл. Выберите :", sys.exc_info()[0])
                    return

            road = wb_data.sheets[0]['A5'].value
            wb_sample.sheets['Исходные данные']['A5'].value = road

            place = wb_data.sheets[0]['A' + str(dataset_begins[i])].value
            wb_sample.sheets['Исходные данные']['A6'].value = place

            place = place.replace('/', '.')  # Слеш заменяю на точку
            place = place.replace('"',
                                  '')  # Удаляю символы кавычек, вопросительный и восклицательный знаки, символ звёздочки, которые могут мешать сохранению файла
            place = place.replace('?', '')
            place = place.replace('!', '')
            place = place.replace('*', '')
            place = place.replace('\n', '')  # Убираю символы переноса строки
            place = place.replace('\t', '')  # Убираю символы табуляции

            # Копирую исходные данные об интенсивности движения
            if type_file == 'Rosautodor_1':
                # my_values = self.wb_data.sheets[0].range('A7:Y' + str(last_row)).options(ndim=2).value
                my_values = wb_data.sheets[0].range(
                    str('A' + str(dataset_begins[i] + 1) + ':Y' + str(dataset_begins[i + 1] - 1))).options(
                    ndim=2).value

            elif type_file == 'Rosautodor_2':
                # my_values = self.wb_data.sheets[0].range('A7:AB' + str(last_row)).options(ndim=2).value
                my_values = wb_data.sheets[0].range(
                    str('A' + str(dataset_begins[i] + 1) + ':AB' + str(dataset_begins[i + 1] - 1))).options(
                    ndim=2).value

            wb_sample.sheets['Исходные данные'].range('A7').value = my_values

            # # Копирую заголовок с описанием типов транспортных средств
            # car_category=wb_data.sheets['Исходные данные'].range('E7:AB7').options(ndim=2).value
            # wb_sample.sheets['Исходные данные'].range('E1:AB1').value = car_category

            # Сохранение под новым названием в специальной папке
            # wb_sample.save(os.path.join('Первичная обработка', 'PRE_' + place + '.xlsx'))
            new_file_name = 'PRE_' + place + '.xlsx'
            self.editor.insert(tk.END, f'\n===== Обрабатываю файл {new_file_name}... =====\n\n')
            wb_sample.save(os.path.join(self.directory_pre, new_file_name))
            wb_sample.close()

            # читаю предобработанный файл
            df = self.checking.open_and_read_file(self.directory_pre, new_file_name, type_file)

            # создаю ID детектора
            if 'ММЗ' in new_file_name:
                detector_id = re.search(r'[A-Z|А-Я]-\d+', new_file_name).group(0).lower()\
                                  .replace('-', '') + '_km' + re.search(r'\d+(?=\+)', new_file_name).group(0) + '_mv'
            else:
                detector_id = re.search(r'[A-Z|А-Я]-\d+', new_file_name).group(0).lower()\
                                  .replace('-', '') + '_km' + re.search(r'\d+(?=\+)', new_file_name).group(0)

            df_total_long = self.checking.make_long(df, self.cur_year, new_file_name, type_file)
            df_main_clear, basic_stats_SSID, basic_stats_intensivnosti = self.checking.fill_gaps_and_remove_outliers(df_total_long, detector_id, type_file)

            # созраняю статистику
            self.bs_SSID.append(basic_stats_SSID)
            self.bs_intensivnosti.append(basic_stats_intensivnosti)

            self.bs_SSID_full = pd.concat(self.bs_SSID)
            self.bs_intensivnosti_full = pd.concat(self.bs_intensivnosti)

            self.bs_SSID_full.to_excel('../out/basic_stats_SSID.xlsx')
            self.bs_intensivnosti_full.to_excel('../out/basic_stats_intensivnosti.xlsx')

            self.checking.plot_graphs(df_main_clear, df_total_long, self.cur_year, new_file_name, freq='d')

        wb_data.app.quit()

    def init_data_import(self, file, rbt_var=None):
        # Функция для вставки исходных данных от детектора в xls-шаблон для последующей обработки.
        # Копирует требуемые диапазоны ячеек в существующий xls-файл.

        # Определяю номер последней строки в исходной таблице
        wb = openpyxl.load_workbook(os.path.join(self.directory_pre, file))
        sheet = wb.active
        columnA = sheet['A']
        last_row = len(sheet['A'])
        #    print(last_row)

        # # Open the Excel program, the default setting: the program is visible, only open without creating a new workbook, and the screen update is turned off
        # app=xw.App(visible=True, add_book=False)
        # app.display_alerts=False
        # app.screen_updating=False

        wb_data = xw.Book(os.path.join(self.directory_pre, file))
        wb_sample = None

        type_file = None
        # Структуру исходных файлов не проверяю, поскольку сделал это уже ранее. Единый формат отчёта
        if rbt_var:
            type_file = self.structure_check(file)

        # Проверка соответствия структуры исходных файлов отключена, поскольку сделал такую проверку ранее отдельно.
        # В текущей версии все файлы от Росавтодора - единого формата.
        if type_file:
            if type_file == 'Rosautodor_1':
                wb_sample = xw.Book('../sample_fda_var1.xlsm')  # 'sample_r1.xlsx'
            elif type_file == 'Rosautodor_2':
                wb_sample = xw.Book('../sample_fda_var2.xlsm')  # 'sample_r.xlsx'
        else:
            # if not type_file:
            tk.messagebox.showwarning(title="ALERT",
                                      message="Вы выбрали не проверять структура файла. Выберите файл шаблона вручную\n"
                                              "Выберите файл шаблона (sample_fda_var1.xlsm) или (sample_fda_var2.xlsm)\n"
                                              "Посмотреть какой именно шаблон подходит к обрабатываемому файлу можно посмотреть в log-файле в корне проекта")
            folder_path_to_sample = filedialog.askopenfile()  # Open a folder selection dialog

            print(folder_path_to_sample)

            if folder_path_to_sample:
                # print(True)
                wb_sample = xw.Book(folder_path_to_sample.name)
            elif not folder_path_to_sample:
                # print(False)
                tk.messagebox.showerror(title="ALERT",
                                        message="Не выбран ни один файл шаблона.\nВыберите и повторите попытку")
                print("Не выбран ни один файл. Выберите :", sys.exc_info()[0])
                return

            # wb_sample = xw.Book('../sample_fda_var2.xlsm')

        # Перевод окон в полноэкранный режим. Макрос прописан в файле шаблона 'sample_r1.xlsm'
        # Текст макроса для VBA (прописывается в Excel):
        # Sub Maximize_Window()
        # Application.WindowState = xlMaximized
        # End Sub

        Maximize = wb_sample.macro('Maximize_Window')
        Maximize()
        time.sleep(5)

        road_name = wb_data.sheets[0]['A6'].value
        road_name = road_name.replace('/', '.')  # Слеш заменяю на точку
        road_name = road_name.replace('"',
                                      '')  # Удаляю символы кавычек, вопросительный и восклицательный знаки, символ звёздочки, которые могут мешать сохранению файла
        road_name = road_name.replace('?', '')
        road_name = road_name.replace('!', '')
        road_name = road_name.replace('*', '')
        road_name = road_name.replace('\n', '')  # Убираю символы переноса строки
        road_name = road_name.replace('\t', '')  # Убираю символы табуляции

        # Копирую данные о названии дороги и местоположении детектора
        my_values = wb_data.sheets[0].range('A5:A6').options(ndim=2).value
        wb_sample.sheets['Исходные данные'].range('A3:A4').value = my_values

        # Сведения о категориях ТС не копируются, поскольку уже заложены в шаблон 'sample_r1.xlsx'

        # Изначально копировал исходные данные об интенсивности движения только по 24*366=8784 строкам
        # Выяснилось, что в отчётах могут быть данные за несколько лет и копировать нужно все данные
        my_values = wb_data.sheets[0].range('A7:AB' + str(last_row)).options(ndim=2).value
        wb_sample.sheets['Исходные данные'].range('A5:AB' + str(last_row - 2)).value = my_values

        # wb_sample.sheets['Итоги'].activate()
        # time.sleep(5)
        # image = pyscreenshot.grab()
        # # print(file)
        # image.save(os.path.join('Графики', filename[0:len(filename) - 5] + '.png'))

        # # Копирую заголовок с описанием типов транспортных средств
        # car_category=wb_data.sheets['Исходные данные'].range('E7:AB7').options(ndim=2).value
        # wb_sample.sheets['Исходные данные'].range('E1:AB1').value = car_category

        #     # Перехожу на вкладку "Итоги" и после этого ожидаю 5 секунд для отрисовки графиков
        # #    wb_sample.sheets["Итоги"].api.Tab.Color = rgb_to_int((146, 208, 80))
        #     wb_sample.sheets['Итоги'].activate()
        #     print('OK')
        #     time.sleep(5)

        # Сохранение под новым названием в специальной папке
        # wb_sample.save(os.path.join('Импортированные данные', 'IN_'+road_name+'.xlsx'))
        # df_long = os.path.join(directory_imp, 'IN_' + road_name + file[4:])
        wb_sample.save(os.path.join(self.directory_imp, 'IN_' + file[4:]))

        # делаю скриншоты сразу
        # wb_data = xw.Book(os.path.join(directory_imp, file))
        wb_sample.sheets['Итоги'].activate()
        # Maximize = wb_data.macro('Maximize_Window')
        # Maximize()
        time.sleep(3)
        image = pyscreenshot.grab()
        # print(file)
        # short_filename = re.search('IN_.+', str(file))[0][3:]
        # Первое совпадение, начинающееся на "флаг" IN_, с отбрасыванием этого флага при формировании имени файла (т.е. в имя файла включаются символы, начиная с четвёртого)
        # print(short_filename)
        # image.save(os.path.join('Графики', file[0:len(short_filename) - 5] + '.png'))
        image.save(os.path.join(self.directory_pic, 'PIC_' + file[3:len(file) - 5] + '.png'))
        wb_data.app.quit()
        # wb_sample.app.quit()

    def screenshots(self, file):
        wb_service = xw.Book('../maximize.xlsm')
        Maximize = wb_service.macro('Maximize_Window')
        Maximize()

        wb_data = xw.Book(os.path.join(self.directory_imp, file))
        wb_data.sheets['Итоги'].activate()
        # Maximize = wb_data.macro('Maximize_Window')
        # Maximize()
        time.sleep(3)
        image = pyscreenshot.grab()
        # print(file)
        # short_filename = re.search('IN_.+', str(file))[0][3:]
        # Первое совпадение, начинающееся на "флаг" IN_, с отбрасыванием этого флага при формировании имени файла (т.е. в имя файла включаются символы, начиная с четвёртого)
        # print(short_filename)
        # image.save(os.path.join('Графики', file[0:len(short_filename) - 5] + '.png'))
        image.save(os.path.join(self.directory_pic, 'PIC_' + file[3:len(file) - 5] + '.png'))
        wb_data.app.quit()


if __name__ == '__main__':
    window = Tk()
    app = Application(window)
    window.mainloop()
