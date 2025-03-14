import datetime as dt
import os
import time

import numpy as np
import pandas as pd
import xlwings as xw
import openpyxl
# from statsmodels.tsa.seasonal import MSTL
# import matplotlib

# matplotlib.use('TkAgg')
from matplotlib import pyplot as plt
# from matplotlib.ticker import MaxNLocator
from matplotlib.dates import MonthLocator, DateFormatter
from matplotlib.gridspec import GridSpec
import matplotlib.ticker as mtick
# plt.switch_backend('agg')
from tqdm import tqdm
import tkinter as tk
from tkinter import messagebox

from logger_setup import setup_logger
logger_FDA, logger_GK = setup_logger()


class Checking:

    def __init__(self, editor, parent):
        self.small_statistics = []
        self.bs_SSID = []
        self.bs_intensivnosti = []
        self.errors_statistics = pd.DataFrame({('Дата', 'Дата'): {},
                                          ('Величина ошибки', 'Логические'): {},
                                          ('Величина ошибки', 'Лишние данные'): {},
                                          ('Величина ошибки', 'Количество мотоциклов'): {},
                                          ('Величина ошибки', 'Некорректные данные Прямое'): {},
                                          ('Величина ошибки', 'Некорректные данные Обратное'): {}})
        self.wb_data = None
        self.column_names = None
        self.time_interval_cond = None
        self.coef_sample = None
        self.parent = parent
        self.editor = editor

    def open_and_read_file(self, directory_pre, file, which_sample):
        try:
            # directory_pre = '../raw_data/ГК/2024/Первичная обработка'
            # which_sample = 'Autodor_3'
            self.wb_data = xw.Book(os.path.join(directory_pre, file))
            wb = openpyxl.load_workbook(os.path.join(directory_pre, file))
            sheet = wb.active
            last_row = len(sheet['A'])
            if which_sample == 'Rosautodor_1':
                my_values = self.wb_data.sheets[0].range('A7:Y' + str(last_row)).options(ndim=2).value
            elif which_sample == 'Rosautodor_2':
                my_values = self.wb_data.sheets[0].range('A7:AB' + str(last_row)).options(ndim=2).value
            elif which_sample == 'Autodor_1':
                my_values = self.wb_data.sheets[0].range('A9:V' + str(last_row)).options(ndim=2).value
            elif which_sample == 'Autodor_2':
                my_values = self.wb_data.sheets[0].range('A9:Y' + str(last_row)).options(ndim=2).value  # включая неопознанные
            elif which_sample == 'Autodor_3':
                my_values = self.wb_data.sheets[0].range('A9:Y' + str(last_row)).options(ndim=2).value  # включая неопознанные

            df = pd.DataFrame(my_values)
            df.fillna(0, inplace=True)

            return df
        except Exception as e:
            print(f"Ошибка при чтении файла {file}: {e}")
            tk.messagebox.showerror(title="ALERT",
                                    message=f"Ошибка при чтении файла {file}: {e}")
            return None

    def make_long(self, df, company, cur_year, file, which_sample):
        self.editor.insert(tk.END, f'Превращаю данные файла {file} в длинный формат...\n')
        self.editor.see(tk.END)

        if company == 'ФДА':
            # удаляю лишние строки в файлах от ФДА: Итого, Среднее и %
            df = df.iloc[:df[df[df.columns[0]] == 'Итого'].index[0]].copy()

        if which_sample == 'Rosautodor_1':
            self.column_names = self.wb_data.sheets['Исходные данные'].range('A3:Y4').value
        elif which_sample == 'Rosautodor_2':
            self.column_names = self.wb_data.sheets['Исходные данные'].range('A3:AB4').value
        elif which_sample in ['Autodor_2', 'Autodor_3']:
            self.column_names = self.wb_data.sheets['Исходные данные'].range('A7:Y8').value
        elif which_sample == 'Autodor_1':
            self.column_names = self.wb_data.sheets['Исходные данные'].range('A7:V8').value

        if company == 'ГК':
            self.wb_data.app.quit()

        self.column_names = pd.DataFrame(self.column_names).ffill(axis=1).ffill(axis=0)

        if company == 'ГК':
            self.column_names.iloc[:, 0] = 'Дата'
            self.column_names.iloc[0, 1:4] = 'Общая интенсивность автомобилей'
            self.column_names.replace(regex={r'^.сего$': 'Итого', '^.рямое$': 'Прямое', '^.братное$': 'Обратное'}, inplace=True)

        df.columns = pd.MultiIndex.from_arrays(self.column_names[:2].values, names=['type_vehicle', 'direction'])
        df[('Дата', 'Дата')] = pd.to_datetime(df[('Дата', 'Дата')], format='%d.%m.%Y %H:%M:%S')
        df.iloc[:, 1:] = df.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

        # принудительно делаем из object - float64
        for col in df.columns[1:]:
            if isinstance(df[col].dtype, object):
                df[col] = df[col].astype(float)

        # Добавляем декабрьские данные для ФДА, потому что там другая логика
        if company in 'ФДА':
            tmp_df = self.__add_previous_december_data(company, cur_year, file, which_sample)
            tmp_df.fillna(0, inplace=True)
            list_dfs = [df, tmp_df]
            if not tmp_df.empty:
                df = pd.concat([i.dropna(axis=1, how='all') for i in list_dfs]).reset_index(drop=True)

        # Достраиваю временной ряд, если в исходных данных полностью отсутствуют дни/часы замеров
        full_time_series = pd.DataFrame({('Дата', 'Дата'):
                                             pd.date_range('2023-12-01 00:59:59', '2025-01-31 23:59:59', freq='1h')})
        full_time_series = pd.to_datetime(full_time_series[('Дата', 'Дата')], format='%d.%m.%Y %H:%M:%S')
        df = full_time_series.to_frame().merge(df, how='left', on=[('Дата', 'Дата')])
        df.fillna(0, inplace=True)
        df.set_index(('Дата', 'Дата'), inplace=True)
        self.time_interval_cond = full_time_series[
            ('2024-01-01 00:59:59' < full_time_series) & (full_time_series < '2024-12-31 23:59:59')]

        max_measured = 8784  # в 2024 году было 366 дней по 24 часа каждый (эту штуку надо автоматизировать)
        fact_measured = sum(
            df[('Общая интенсивность автомобилей', 'Итого')][self.time_interval_cond].replace(0, np.nan).notna())
        fullness = fact_measured / max_measured
        self.editor.insert(tk.END, f'Максимальное количество замеров: {max_measured}\n'
                                   f'Фактическое количество замеров: {fact_measured}\n'
                                   f'Полнота данных: {fullness:,.1%}\n')
        self.editor.see(tk.END)

        # фигово, что приходится создавать целую переменную для такого, но уже ничего не поделать...
        self.small_statistics = [max_measured, fact_measured, fullness]

        checker_data, self.errors_statistics = self.__check_correct_data(df, which_sample)

        df_long = df.melt(ignore_index=False, value_name='Количество').reset_index().rename(
            columns={('Дата', 'Дата'): 'Дата', 'variable_0': 'type_vehicle', 'variable_1': 'direction'})
        checker_data_long = checker_data.melt(ignore_index=False, col_level=1, value_vars=['Обратное', 'Прямое', 'Итого'],
                                              var_name='direction', value_name='Корректность').reset_index().rename(
            columns={('Дата', 'Дата'): "Дата"})

        df_total_long = df_long.merge(checker_data_long, how='left', on=['Дата',
                                                                         'direction'])  # .rename(columns={'Значение_x': 'Количество', 'Значение_y': 'Корректность'})

        # тут надо повнимательнее! может быть переписать этот кусочек... потому что эти два показателя залезают
        # в данные только при первом варианте ФДА Росавтодор (загрузка и скорость)
        if company == 'ФДА':
            df_total_long = df_total_long[
                (df_total_long.type_vehicle != 'Загрузка, %') & (df_total_long.type_vehicle != 'Скорость, км/ч')]

        df_total_long.set_index('Дата', inplace=True)
        # df_total_long.
        return df_total_long

    def __add_previous_december_data(self, company, cur_year, file, which_sample):
        previous_year_files = os.listdir(f'../raw_data/{company}/{str(int(cur_year) - 1)}/Первичная обработка')
        wb_prev_yr_file = pd.DataFrame()
        try:
            for i in previous_year_files:
                # print(i)
                if file == i:
                    self.parent.update_idletasks()
                    time.sleep(2)
                    self.editor.insert(tk.END, f'Файл за предыдущий год найден! {i} \n'
                                               f'Добавляю к текущему данные за декабрь предыдущего, {str(int(cur_year) - 1)}...\n')
                    print('Файл за предыдущий год найден! ' + i)
                    self.editor.see(tk.END)
                    # tk.messagebox.showerror(title="ALERT",
                    #                         message=f"Файл за предыдущий год найден! {i}")
                    if which_sample == 'Rosautodor_1':
                        wb_prev_yr_file = pd.read_excel(
                            os.path.join(f'../raw_data/{company}/{str(int(cur_year) - 1)}/Первичная обработка/', file)).iloc[:,
                                          0:-4]
                    elif which_sample == 'Rosautodor_2':
                        wb_prev_yr_file = pd.read_excel(
                            os.path.join(f'../raw_data/{company}/{str(int(cur_year) - 1)}/Первичная обработка/', file)).iloc[:,
                                          0:-4]

                    wb_prev_yr_file.columns = pd.MultiIndex.from_arrays(self.column_names[:2].values,
                                                                        names=['type_vehicle', 'direction'])
                    wb_prev_yr_file.set_index(wb_prev_yr_file.columns[0], inplace=True)
                    wb_prev_yr_file.drop(wb_prev_yr_file.head(5).index, inplace=True)
                    wb_prev_yr_file.drop(index=['Итого', 'Среднее', '%'], axis=0, inplace=True)
                    wb_prev_yr_file.reset_index(inplace=True)
                    wb_prev_yr_file[('Дата', 'Дата')] = pd.to_datetime(wb_prev_yr_file[('Дата', 'Дата')],
                                                                       format='%d.%m.%Y %H:%M:%S')
                    # wb_prev_yr_file.dtypes

                    wb_prev_yr_file = wb_prev_yr_file[wb_prev_yr_file[('Дата', 'Дата')].dt.month == 12]
                    wb_prev_yr_file.iloc[:, 1:] = wb_prev_yr_file.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

                    # принудительно делаем из object - float64
                    for col in wb_prev_yr_file.columns[1:]:
                        if isinstance(wb_prev_yr_file[col].dtype, object):
                            wb_prev_yr_file[col] = wb_prev_yr_file[col].astype(float)
                else:
                    pass
            return wb_prev_yr_file
        except Exception as e:
            self.editor.insert(tk.END, f'Не удалось найти файл {file} за прошлый год: {e}. \n')
            self.editor.see(tk.END)
            print(f"Не удалось найти файл {file} за прошлый год: {e}")
            tk.messagebox.showerror(title="ALERT",
                                    message=f"Не удалось найти файл {file} за прошлый год: {e}")
            return wb_prev_yr_file

    def __check_correct_data(self, df, which_sample):
        self.editor.insert(tk.END, f'Проверяю корректность данных и считаю статистику... \n')
        self.editor.see(tk.END)
        sum_non_passenger_cars = pd.DataFrame({('Дата', 'Дата'): {},
                                               ('Сумма не-легковых', 'Итого'): {},
                                               ('Сумма не-легковых', 'Прямое'): {},
                                               ('Сумма не-легковых', 'Обратное'): {}})
        sum_non_passenger_cars['Дата', 'Дата'] = df.index
        sum_non_passenger_cars.set_index(('Дата', 'Дата'), inplace=True)
        try:
            if which_sample in ['Rosautodor_1', 'Rosautodor_2']:

                # sum_non_passenger_cars.columns = pd.MultiIndex.from_arrays(column_names.values, names=['type_vehicle', 'direction'])
                sum_non_passenger_cars['Сумма не-легковых', 'Обратное'] = df[[('Малые груз. (6-9 м)', 'Обратное'),
                                                                              ('Грузовые (9-13 м)', 'Обратное'),
                                                                              ('Груз. большие (13-22 м)', 'Обратное'),
                                                                              ('Автопоезда (22-30 м)', 'Обратное'),
                                                                              ('Автобусы', 'Обратное')]].sum(axis=1)
                sum_non_passenger_cars['Сумма не-легковых', 'Прямое'] = df[[('Малые груз. (6-9 м)', 'Прямое'),
                                                                            ('Грузовые (9-13 м)', 'Прямое'),
                                                                            ('Груз. большие (13-22 м)', 'Прямое'),
                                                                            ('Автопоезда (22-30 м)', 'Прямое'),
                                                                            ('Автобусы', 'Прямое')]].sum(axis=1)
                sum_non_passenger_cars['Сумма не-легковых', 'Итого'] = df[[('Малые груз. (6-9 м)', 'Итого'),
                                                                           ('Грузовые (9-13 м)', 'Итого'),
                                                                           ('Груз. большие (13-22 м)', 'Итого'),
                                                                           ('Автопоезда (22-30 м)', 'Итого'),
                                                                           ('Автобусы', 'Итого')]].sum(axis=1)

            elif which_sample == 'Autodor_1':
                sum_non_passenger_cars['Сумма не-легковых', 'Обратное'] = df[[('микроавтобусы, малые грузовики', 'Обратное'),
                                                                              ('одиночные АТС, автобусы', 'Обратное'),
                                                                              ('автопоезда до 13 м', 'Обратное'),
                                                                              ('автопоезда 13..18  м', 'Обратное'),
                                                                              ('длинные автопоезда свыше 18 м', 'Обратное')]].sum(axis=1)
                sum_non_passenger_cars['Сумма не-легковых', 'Прямое'] = df[[('микроавтобусы, малые грузовики', 'Прямое'),
                                                                            ('одиночные АТС, автобусы', 'Прямое'),
                                                                            ('автопоезда до 13 м', 'Прямое'),
                                                                            ('автопоезда 13..18  м', 'Прямое'),
                                                                            ('длинные автопоезда свыше 18 м', 'Прямое')]].sum(axis=1)
                sum_non_passenger_cars['Сумма не-легковых', 'Итого'] = df[[('микроавтобусы, малые грузовики', 'Итого'),
                                                                           ('одиночные АТС, автобусы', 'Итого'),
                                                                           ('автопоезда до 13 м', 'Итого'),
                                                                           ('автопоезда 13..18  м', 'Итого'),
                                                                           ('длинные автопоезда свыше 18 м', 'Итого')]].sum(axis=1)
            elif which_sample == 'Autodor_2':
                sum_non_passenger_cars['Сумма не-легковых', 'Обратное'] = df[[('микроавтобусы, малые грузовые автомобили (6-9 м)', 'Обратное'),
                                                                              ('грузовые автомобили (9-11 м)', 'Обратное'),
                                                                              ('автобусы (11-13 м)', 'Обратное'),
                                                                              ('грузовые большие автомобили, автопоезда (13-18 м)', 'Обратное'),
                                                                              ('длинные автопоезда (> 18 м)', 'Обратное')]].sum(axis=1)
                sum_non_passenger_cars['Сумма не-легковых', 'Прямое'] = df[[('микроавтобусы, малые грузовые автомобили (6-9 м)', 'Прямое'),
                                                                            ('грузовые автомобили (9-11 м)', 'Прямое'),
                                                                            ('автобусы (11-13 м)', 'Прямое'),
                                                                            ('грузовые большие автомобили, автопоезда (13-18 м)', 'Прямое'),
                                                                            ('длинные автопоезда (> 18 м)', 'Прямое')]].sum(axis=1)
                sum_non_passenger_cars['Сумма не-легковых', 'Итого'] = df[[('микроавтобусы, малые грузовые автомобили (6-9 м)', 'Итого'),
                                                                           ('грузовые автомобили (9-11 м)', 'Итого'),
                                                                           ('автобусы (11-13 м)', 'Итого'),
                                                                           ('грузовые большие автомобили, автопоезда (13-18 м)', 'Итого'),
                                                                           ('длинные автопоезда (> 18 м)', 'Итого')]].sum(axis=1)
            elif which_sample == 'Autodor_3':
                sum_non_passenger_cars['Сумма не-легковых', 'Обратное'] = df[[('малые грузовые автомобили до 5 тонн (6-9 м)', 'Обратное'),
                                                                              ('грузовые автомобили 5-12 тонн (9-11 м)', 'Обратное'),
                                                                              ('автобусы (11-13 м)', 'Обратное'),
                                                                              ('грузовые большие автомобили 12-20 тонн (13-22 м)', 'Обратное'),
                                                                              ('автопоезда более 20 тонн (22-30 м)', 'Обратное')]].sum(axis=1)
                sum_non_passenger_cars['Сумма не-легковых', 'Прямое'] = df[[('малые грузовые автомобили до 5 тонн (6-9 м)', 'Прямое'),
                                                                            ('грузовые автомобили 5-12 тонн (9-11 м)', 'Прямое'),
                                                                            ('автобусы (11-13 м)', 'Прямое'),
                                                                            ('грузовые большие автомобили 12-20 тонн (13-22 м)', 'Прямое'),
                                                                            ('автопоезда более 20 тонн (22-30 м)', 'Прямое')]].sum(axis=1)
                sum_non_passenger_cars['Сумма не-легковых', 'Итого'] = df[[('малые грузовые автомобили до 5 тонн (6-9 м)', 'Итого'),
                                                                           ('грузовые автомобили 5-12 тонн (9-11 м)', 'Итого'),
                                                                           ('автобусы (11-13 м)', 'Итого'),
                                                                           ('грузовые большие автомобили 12-20 тонн (13-22 м)', 'Итого'),
                                                                           ('автопоезда более 20 тонн (22-30 м)', 'Итого')]].sum(axis=1)

        except KeyError as e:
            print('Плашка')
            tk.messagebox.showerror(title="ALERT",
                                    message=f"Не удалось посчитать количество нелегковых автомобилей {e}.\n"
                                            f"Видимо, названия в столбцах исходных данных изменились."
                                            f"Добавьте или измените названия в df в check_correct.py (строки 228-242)")
            return

        # error_amount_direction - Величина ошибки при подсчёте сумм по направлениям (лишние данные) если Итого оказалось больше/меньше чем сумма прямого и обратного
        error_amount_direction = pd.DataFrame({('Дата', 'Дата'): {},
                                               ('Величина ошибки', 'Прямое'): {},
                                               ('Величина ошибки', 'Обратное'): {}})
        error_amount_direction['Дата', 'Дата'] = df.index
        error_amount_direction.set_index(('Дата', 'Дата'), inplace=True)

        # Считаю логические ошибки
        logical_errors_GK = pd.DataFrame({('Дата', 'Дата'): {},
                                          ('Логические ошибки', 'Общая интенсивность автомобилей'): {},
                                          ('Логические ошибки', 'легковые'): {},
                                          ('Логические ошибки', 'микроавтобусы, малые грузовики'): {},
                                          ('Логические ошибки', 'одиночные АТС, автобусы'): {},
                                          ('Логические ошибки', 'автопоезда до 13 м'): {},
                                          ('Логические ошибки', 'автопоезда 13..18  м'): {},
                                          ('Логические ошибки', 'длинные автопоезда свыше 18 м'): {},
                                          ('Логические ошибки', 'легковые автомобили (до 6 м)'): {},
                                          ('Логические ошибки', 'микроавтобусы, малые грузовые автомобили (6-9 м)'): {},
                                          ('Логические ошибки', 'грузовые автомобили (9-11 м)'): {},
                                          ('Логические ошибки', 'автобусы (11-13 м)'): {},
                                          ('Логические ошибки', 'грузовые большие автомобили, автопоезда (13-18 м)'): {},
                                          ('Логические ошибки', 'длинные автопоезда (> 18 м)'): {},
                                          ('Логические ошибки', 'неопознаные тс'): {},
                                          # ('Логические ошибки', 'легковые автомобили (до 6 м)'): {},
                                          ('Логические ошибки', 'малые грузовые автомобили до 5 тонн (6-9 м)'): {},
                                          ('Логические ошибки', 'грузовые автомобили 5-12 тонн (9-11 м)'): {},
                                          # ('Логические ошибки', 'автобусы (11-13 м)'): {},
                                          ('Логические ошибки', 'грузовые большие автомобили 12-20 тонн (13-22 м)'): {},
                                          ('Логические ошибки', 'автопоезда более 20 тонн (22-30 м)'): {},
                                          # ('Логические ошибки', 'неопознаные тс'): {},
                                          })

        logical_errors = pd.DataFrame({('Дата', 'Дата'): {},
                                       ('Логические ошибки', 'Общая интенсивность автомобилей'): {},
                                       ('Логические ошибки', 'Легковые (до 4.5 м)'): {},
                                       ('Логические ошибки', 'Легковые большие (4-6 м)'): {},
                                       ('Логические ошибки', 'Легковые (до 6 м)'): {},
                                       ('Логические ошибки', 'Малые груз. (6-9 м)'): {},
                                       ('Логические ошибки', 'Грузовые (9-13 м)'): {},
                                       ('Логические ошибки', 'Груз. большие (13-22 м)'): {},
                                       ('Логические ошибки', 'Автопоезда (22-30 м)'): {},
                                       ('Логические ошибки', 'Автобусы'): {},
                                       ('Логические ошибки', 'Мотоциклы'): {},
                                       ('Логические ошибки', 'Прямое'): {},
                                       ('Логические ошибки', 'Обратное'): {},})

        logical_errors_GK['Дата', 'Дата'] = df.index
        logical_errors_GK.set_index(('Дата', 'Дата'), inplace=True)

        logical_errors['Дата', 'Дата'] = df.index
        logical_errors.set_index(('Дата', 'Дата'), inplace=True)

        if which_sample in ['Rosautodor_1', 'Rosautodor_2']:
            logical_errors['Логические ошибки',
                           'Общая интенсивность автомобилей'] = np.where((df['Общая интенсивность автомобилей', 'Итого'] -
                                                                          df['Общая интенсивность автомобилей', 'Прямое'] -
                                                                          df[
                                                                              'Общая интенсивность автомобилей', 'Обратное']) == 0,
                                                                         0, 1)

            logical_errors['Логические ошибки',
                           'Малые груз. (6-9 м)'] = np.where((df['Малые груз. (6-9 м)', 'Итого'] -
                                                              df['Малые груз. (6-9 м)', 'Прямое'] -
                                                              df['Малые груз. (6-9 м)', 'Обратное']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Грузовые (9-13 м)'] = np.where((df['Грузовые (9-13 м)', 'Итого'] -
                                                            df['Грузовые (9-13 м)', 'Прямое'] -
                                                            df['Грузовые (9-13 м)', 'Обратное']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Груз. большие (13-22 м)'] = np.where((df['Груз. большие (13-22 м)', 'Итого'] -
                                                                  df['Груз. большие (13-22 м)', 'Прямое'] -
                                                                  df['Груз. большие (13-22 м)', 'Обратное']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Автопоезда (22-30 м)'] = np.where((df['Автопоезда (22-30 м)', 'Итого'] -
                                                               df['Автопоезда (22-30 м)', 'Прямое'] -
                                                               df['Автопоезда (22-30 м)', 'Обратное']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Автобусы'] = np.where((df['Автобусы', 'Итого'] -
                                                   df['Автобусы', 'Прямое'] -
                                                   df['Автобусы', 'Обратное']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Мотоциклы'] = np.where((df['Мотоциклы', 'Итого'] -
                                                    df['Мотоциклы', 'Прямое'] -
                                                    df['Мотоциклы', 'Обратное']) == 0, 0, 1)

        if which_sample == 'Rosautodor_1':
            logical_errors['Логические ошибки',
                           'Легковые (до 6 м)'] = np.where((df['Легковые (до 6 м)', 'Итого'] -
                                                            df['Легковые (до 6 м)', 'Прямое'] -
                                                            df['Легковые (до 6 м)', 'Обратное']) == 0, 0, 1)

            logical_errors['Логические ошибки',
                           'Прямое'] = np.where((df['Общая интенсивность автомобилей', 'Прямое'] -
                                                 df['Легковые (до 6 м)', 'Прямое'] -
                                                 df['Малые груз. (6-9 м)', 'Прямое'] -
                                                 df['Грузовые (9-13 м)', 'Прямое'] -
                                                 df['Груз. большие (13-22 м)', 'Прямое'] -
                                                 df['Автопоезда (22-30 м)', 'Прямое'] -
                                                 df['Автобусы', 'Прямое'] -
                                                 df['Мотоциклы', 'Прямое']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Обратное'] = np.where((df['Общая интенсивность автомобилей', 'Обратное'] -
                                                   df['Легковые (до 6 м)', 'Обратное'] -
                                                   df['Малые груз. (6-9 м)', 'Обратное'] -
                                                   df['Грузовые (9-13 м)', 'Обратное'] -
                                                   df['Груз. большие (13-22 м)', 'Обратное'] -
                                                   df['Автопоезда (22-30 м)', 'Обратное'] -
                                                   df['Автобусы', 'Обратное'] -
                                                   df['Мотоциклы', 'Обратное']) == 0, 0, 1)

            error_amount_direction['Величина ошибки', 'Прямое'] = df['Общая интенсивность автомобилей', 'Прямое'] - \
                                                                  df['Легковые (до 6 м)', 'Прямое'] - \
                                                                  df['Малые груз. (6-9 м)', 'Прямое'] - \
                                                                  df['Грузовые (9-13 м)', 'Прямое'] - \
                                                                  df['Груз. большие (13-22 м)', 'Прямое'] - \
                                                                  df['Автопоезда (22-30 м)', 'Прямое'] - \
                                                                  df['Автобусы', 'Прямое'] - \
                                                                  df['Мотоциклы', 'Прямое']
            error_amount_direction['Величина ошибки', 'Обратное'] = df['Общая интенсивность автомобилей', 'Обратное'] - \
                                                                    df['Легковые (до 6 м)', 'Обратное'] - \
                                                                    df['Малые груз. (6-9 м)', 'Обратное'] - \
                                                                    df['Грузовые (9-13 м)', 'Обратное'] - \
                                                                    df['Груз. большие (13-22 м)', 'Обратное'] - \
                                                                    df['Автопоезда (22-30 м)', 'Обратное'] - \
                                                                    df['Автобусы', 'Обратное'] - \
                                                                    df['Мотоциклы', 'Обратное']
            df_tmp_to = df['Легковые (до 6 м)', 'Прямое']
            df_tmp_back = df['Легковые (до 6 м)', 'Обратное']

        elif which_sample == 'Rosautodor_2':
            logical_errors['Логические ошибки',
                           'Легковые (до 4.5 м)'] = np.where((df['Легковые (до 4.5 м)', 'Итого'] -
                                                            df['Легковые (до 4.5 м)', 'Прямое'] -
                                                            df['Легковые (до 4.5 м)', 'Обратное']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Легковые большие (4-6 м)'] = np.where((df['Легковые большие (4-6 м)', 'Итого'] -
                                                            df['Легковые большие (4-6 м)', 'Прямое'] -
                                                            df['Легковые большие (4-6 м)', 'Обратное']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Прямое'] = np.where((df['Общая интенсивность автомобилей', 'Прямое'] -
                                                 df['Легковые (до 4.5 м)', 'Прямое'] -
                                                 df['Легковые большие (4-6 м)', 'Прямое'] -
                                                 df['Малые груз. (6-9 м)', 'Прямое'] -
                                                 df['Грузовые (9-13 м)', 'Прямое'] -
                                                 df['Груз. большие (13-22 м)', 'Прямое'] -
                                                 df['Автопоезда (22-30 м)', 'Прямое'] -
                                                 df['Автобусы', 'Прямое'] -
                                                 df['Мотоциклы', 'Прямое']) == 0, 0, 1)
            logical_errors['Логические ошибки',
                           'Обратное'] = np.where((df['Общая интенсивность автомобилей', 'Обратное'] -
                                                   df['Легковые (до 4.5 м)', 'Обратное'] -
                                                   df['Легковые большие (4-6 м)', 'Обратное'] -
                                                   df['Малые груз. (6-9 м)', 'Обратное'] -
                                                   df['Грузовые (9-13 м)', 'Обратное'] -
                                                   df['Груз. большие (13-22 м)', 'Обратное'] -
                                                   df['Автопоезда (22-30 м)', 'Обратное'] -
                                                   df['Автобусы', 'Обратное'] -
                                                   df['Мотоциклы', 'Обратное']) == 0, 0, 1)

            error_amount_direction['Величина ошибки', 'Прямое'] = df['Общая интенсивность автомобилей', 'Прямое'] - \
                                                                  df['Легковые (до 4.5 м)', 'Прямое'] - \
                                                                  df['Легковые большие (4-6 м)', 'Прямое'] - \
                                                                  df['Малые груз. (6-9 м)', 'Прямое'] - \
                                                                  df['Грузовые (9-13 м)', 'Прямое'] - \
                                                                  df['Груз. большие (13-22 м)', 'Прямое'] - \
                                                                  df['Автопоезда (22-30 м)', 'Прямое'] - \
                                                                  df['Автобусы', 'Прямое'] - \
                                                                  df['Мотоциклы', 'Прямое']
            error_amount_direction['Величина ошибки', 'Обратное'] = df['Общая интенсивность автомобилей', 'Обратное'] - \
                                                                    df['Легковые (до 4.5 м)', 'Обратное'] - \
                                                                    df['Легковые большие (4-6 м)', 'Обратное'] - \
                                                                    df['Малые груз. (6-9 м)', 'Обратное'] - \
                                                                    df['Грузовые (9-13 м)', 'Обратное'] - \
                                                                    df['Груз. большие (13-22 м)', 'Обратное'] - \
                                                                    df['Автопоезда (22-30 м)', 'Обратное'] - \
                                                                    df['Автобусы', 'Обратное'] - \
                                                                    df['Мотоциклы', 'Обратное']
            df_tmp_to = df['Легковые (до 4.5 м)', 'Прямое'] + df['Легковые большие (4-6 м)', 'Прямое']
            df_tmp_back = df['Легковые (до 4.5 м)', 'Обратное'] + df['Легковые большие (4-6 м)', 'Обратное']

        elif which_sample == 'Autodor_1':
            logical_errors_GK['Логические ошибки',
                           'Общая интенсивность автомобилей'] = np.where((df['Общая интенсивность автомобилей',
                                                                             'Итого'] -
                                                                          df['Общая интенсивность автомобилей',
                                                                             'Прямое'] -
                                                                          df['Общая интенсивность автомобилей',
                                                                             'Обратное']) == 0, 0, 1)

            logical_errors_GK['Логические ошибки',
                           'легковые'] = np.where((df['легковые', 'Итого'] -
                                                              df['легковые', 'Прямое'] -
                                                              df['легковые', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'микроавтобусы, малые грузовики'] = np.where((df['микроавтобусы, малые грузовики', 'Итого'] -
                                                            df['микроавтобусы, малые грузовики', 'Прямое'] -
                                                            df['микроавтобусы, малые грузовики', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'одиночные АТС, автобусы'] = np.where((df['одиночные АТС, автобусы', 'Итого'] -
                                                                  df['одиночные АТС, автобусы', 'Прямое'] -
                                                                  df['одиночные АТС, автобусы', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'автопоезда до 13 м'] = np.where((df['автопоезда до 13 м', 'Итого'] -
                                                               df['автопоезда до 13 м', 'Прямое'] -
                                                               df['автопоезда до 13 м', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'автопоезда 13..18  м'] = np.where((df['автопоезда 13..18  м', 'Итого'] -
                                                   df['автопоезда 13..18  м', 'Прямое'] -
                                                   df['автопоезда 13..18  м', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'длинные автопоезда свыше 18 м'] = np.where((df['длинные автопоезда свыше 18 м', 'Итого'] -
                                                    df['длинные автопоезда свыше 18 м', 'Прямое'] -
                                                    df['длинные автопоезда свыше 18 м', 'Обратное']) == 0, 0, 1)

            error_amount_direction['Величина ошибки', 'Прямое'] = df['Общая интенсивность автомобилей', 'Прямое'] - \
                                                                  df['легковые', 'Прямое'] - \
                                                                  df['микроавтобусы, малые грузовики', 'Прямое'] - \
                                                                  df['одиночные АТС, автобусы', 'Прямое'] - \
                                                                  df['автопоезда до 13 м', 'Прямое'] - \
                                                                  df['автопоезда 13..18  м', 'Прямое'] - \
                                                                  df['длинные автопоезда свыше 18 м', 'Прямое']
            error_amount_direction['Величина ошибки', 'Обратное'] = df['Общая интенсивность автомобилей', 'Обратное'] - \
                                                                    df['легковые', 'Обратное'] - \
                                                                    df['микроавтобусы, малые грузовики', 'Обратное'] - \
                                                                    df['одиночные АТС, автобусы', 'Обратное'] - \
                                                                    df['автопоезда до 13 м', 'Обратное'] - \
                                                                    df['автопоезда 13..18  м', 'Обратное'] - \
                                                                    df['длинные автопоезда свыше 18 м', 'Обратное']
            df_tmp_to = df['легковые', 'Прямое']
            df_tmp_back = df['легковые', 'Обратное']

        elif which_sample == 'Autodor_2':
            logical_errors_GK['Логические ошибки',
                           'Общая интенсивность автомобилей'] = np.where((df['Общая интенсивность автомобилей',
                                                                             'Итого'] -
                                                                          df['Общая интенсивность автомобилей',
                                                                             'Прямое'] -
                                                                          df['Общая интенсивность автомобилей',
                                                                             'Обратное']) == 0, 0, 1)

            logical_errors_GK['Логические ошибки',
                           'легковые автомобили (до 6 м)'] = np.where((df['легковые автомобили (до 6 м)', 'Итого'] -
                                                              df['легковые автомобили (до 6 м)', 'Прямое'] -
                                                              df['легковые автомобили (до 6 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'микроавтобусы, малые грузовые автомобили (6-9 м)'] = np.where((df['микроавтобусы, малые грузовые автомобили (6-9 м)', 'Итого'] -
                                                            df['микроавтобусы, малые грузовые автомобили (6-9 м)', 'Прямое'] -
                                                            df['микроавтобусы, малые грузовые автомобили (6-9 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'грузовые автомобили (9-11 м)'] = np.where((df['грузовые автомобили (9-11 м)', 'Итого'] -
                                                                  df['грузовые автомобили (9-11 м)', 'Прямое'] -
                                                                  df['грузовые автомобили (9-11 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'автобусы (11-13 м)'] = np.where((df['автобусы (11-13 м)', 'Итого'] -
                                                               df['автобусы (11-13 м)', 'Прямое'] -
                                                               df['автобусы (11-13 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'грузовые большие автомобили, автопоезда (13-18 м)'] = np.where((df['грузовые большие автомобили, автопоезда (13-18 м)', 'Итого'] -
                                                   df['грузовые большие автомобили, автопоезда (13-18 м)', 'Прямое'] -
                                                   df['грузовые большие автомобили, автопоезда (13-18 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'длинные автопоезда (> 18 м)'] = np.where((df['длинные автопоезда (> 18 м)', 'Итого'] -
                                                    df['длинные автопоезда (> 18 м)', 'Прямое'] -
                                                    df['длинные автопоезда (> 18 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'неопознаные тс'] = np.where((df['неопознаные тс', 'Итого'] -
                                                    df['неопознаные тс', 'Прямое'] -
                                                    df['неопознаные тс', 'Обратное']) == 0, 0, 1)

            error_amount_direction['Величина ошибки', 'Прямое'] = df['Общая интенсивность автомобилей', 'Прямое'] - \
                                                                  df['легковые автомобили (до 6 м)', 'Прямое'] - \
                                                                  df['микроавтобусы, малые грузовые автомобили (6-9 м)', 'Прямое'] - \
                                                                  df['грузовые автомобили (9-11 м)', 'Прямое'] - \
                                                                  df['автобусы (11-13 м)', 'Прямое'] - \
                                                                  df['грузовые большие автомобили, автопоезда (13-18 м)', 'Прямое'] - \
                                                                  df['длинные автопоезда (> 18 м)', 'Прямое'] - \
                                                                  df['неопознаные тс', 'Прямое']
            error_amount_direction['Величина ошибки', 'Обратное'] = df['Общая интенсивность автомобилей', 'Обратное'] - \
                                                                    df['легковые автомобили (до 6 м)', 'Обратное'] - \
                                                                    df['микроавтобусы, малые грузовые автомобили (6-9 м)', 'Обратное'] - \
                                                                    df['грузовые автомобили (9-11 м)', 'Обратное'] - \
                                                                    df['автобусы (11-13 м)', 'Обратное'] - \
                                                                    df['грузовые большие автомобили, автопоезда (13-18 м)', 'Обратное'] - \
                                                                    df['длинные автопоезда (> 18 м)', 'Обратное'] - \
                                                                    df['неопознаные тс', 'Обратное']
            df_tmp_to = df['легковые автомобили (до 6 м)', 'Прямое']
            df_tmp_back = df['легковые автомобили (до 6 м)', 'Обратное']

        elif which_sample == 'Autodor_3':
            logical_errors_GK['Логические ошибки',
                           'Общая интенсивность автомобилей'] = np.where((df['Общая интенсивность автомобилей',
                                                                             'Итого'] -
                                                                          df['Общая интенсивность автомобилей',
                                                                             'Прямое'] -
                                                                          df['Общая интенсивность автомобилей',
                                                                             'Обратное']) == 0, 0, 1)

            logical_errors_GK['Логические ошибки',
                           'легковые автомобили (до 6 м)'] = np.where((df['легковые автомобили (до 6 м)', 'Итого'] -
                                                              df['легковые автомобили (до 6 м)', 'Прямое'] -
                                                              df['легковые автомобили (до 6 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'малые грузовые автомобили до 5 тонн (6-9 м)'] = np.where((df['малые грузовые автомобили до 5 тонн (6-9 м)', 'Итого'] -
                                                            df['малые грузовые автомобили до 5 тонн (6-9 м)', 'Прямое'] -
                                                            df['малые грузовые автомобили до 5 тонн (6-9 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'грузовые автомобили 5-12 тонн (9-11 м)'] = np.where((df['грузовые автомобили 5-12 тонн (9-11 м)', 'Итого'] -
                                                                  df['грузовые автомобили 5-12 тонн (9-11 м)', 'Прямое'] -
                                                                  df['грузовые автомобили 5-12 тонн (9-11 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'автобусы (11-13 м)'] = np.where((df['автобусы (11-13 м)', 'Итого'] -
                                                               df['автобусы (11-13 м)', 'Прямое'] -
                                                               df['автобусы (11-13 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'грузовые большие автомобили 12-20 тонн (13-22 м)'] = np.where((df['грузовые большие автомобили 12-20 тонн (13-22 м)', 'Итого'] -
                                                   df['грузовые большие автомобили 12-20 тонн (13-22 м)', 'Прямое'] -
                                                   df['грузовые большие автомобили 12-20 тонн (13-22 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'автопоезда более 20 тонн (22-30 м)'] = np.where((df['автопоезда более 20 тонн (22-30 м)', 'Итого'] -
                                                    df['автопоезда более 20 тонн (22-30 м)', 'Прямое'] -
                                                    df['автопоезда более 20 тонн (22-30 м)', 'Обратное']) == 0, 0, 1)
            logical_errors_GK['Логические ошибки',
                           'неопознаные тс'] = np.where((df['неопознаные тс', 'Итого'] -
                                                    df['неопознаные тс', 'Прямое'] -
                                                    df['неопознаные тс', 'Обратное']) == 0, 0, 1)

            error_amount_direction['Величина ошибки', 'Прямое'] = df['Общая интенсивность автомобилей', 'Прямое'] - \
                                                                  df['легковые автомобили (до 6 м)', 'Прямое'] - \
                                                                  df['малые грузовые автомобили до 5 тонн (6-9 м)', 'Прямое'] - \
                                                                  df['грузовые автомобили 5-12 тонн (9-11 м)', 'Прямое'] - \
                                                                  df['автобусы (11-13 м)', 'Прямое'] - \
                                                                  df['грузовые большие автомобили 12-20 тонн (13-22 м)', 'Прямое'] - \
                                                                  df['автопоезда более 20 тонн (22-30 м)', 'Прямое'] - \
                                                                  df['неопознаные тс', 'Прямое']
            error_amount_direction['Величина ошибки', 'Обратное'] = df['Общая интенсивность автомобилей', 'Обратное'] - \
                                                                    df['легковые автомобили (до 6 м)', 'Обратное'] - \
                                                                    df['малые грузовые автомобили до 5 тонн (6-9 м)', 'Обратное'] - \
                                                                    df['грузовые автомобили 5-12 тонн (9-11 м)', 'Обратное'] - \
                                                                    df['автобусы (11-13 м)', 'Обратное'] - \
                                                                    df['грузовые большие автомобили 12-20 тонн (13-22 м)', 'Обратное'] - \
                                                                    df['автопоезда более 20 тонн (22-30 м)', 'Обратное'] - \
                                                                    df['неопознаные тс', 'Обратное']
            df_tmp_to = df['легковые автомобили (до 6 м)', 'Прямое']
            df_tmp_back = df['легковые автомобили (до 6 м)', 'Обратное']

        # Количество логических ошибок в данных
        baz = logical_errors.sum(axis=0).sum()
        # Величина ошибки между суммарным потоком и суммой по группам
        # ("лишние" данные по направлениям, деленное на Общую интенсивность авто по направлениям)
        denominator = (df['Общая интенсивность автомобилей', 'Прямое'] +
                       df['Общая интенсивность автомобилей', 'Обратное']).sum()
        if denominator != 0:
            foo = error_amount_direction.sum(axis=0).sum() / denominator
        else:
            foo = 0  # или какое-то другое значение по умолчанию, например, 0

        if which_sample in ['Rosautodor_1', 'Rosautodor_2']:
            # Кол-во зарегистрированных мотоциклов в исходных данных
            bar = (df['Мотоциклы', 'Прямое'] + df['Мотоциклы', 'Обратное']).sum()

        checker_data = pd.DataFrame({('Дата', 'Дата'): {},
                                     ('Величина ошибки', 'Прямое'): {},
                                     ('Величина ошибки', 'Обратное'): {},
                                     ('Величина ошибки', 'Итого'): {}})
        checker_data['Дата', 'Дата'] = df.index
        checker_data.set_index(('Дата', 'Дата'), inplace=True)

        checker_data['Величина ошибки', 'Прямое'] = np.where(
            (df['Общая интенсивность автомобилей', 'Прямое'].groupby(df.index.floor('D')).transform('sum') < 0.25 *
             df['Общая интенсивность автомобилей', 'Обратное'].groupby(df.index.floor('D')).transform('sum')) |
            (df.sum(axis=1) == 0) |  # .loc[(df.index > '2024-01-11 11:59:59') & (df.index < '15.01.2024 16:59:59')]
            (sum_non_passenger_cars['Сумма не-легковых', 'Прямое'] > 10 * df_tmp_to),
            'Данные НЕкорректны', 'Данные корректны')

        checker_data['Величина ошибки', 'Обратное'] = np.where(
            (df['Общая интенсивность автомобилей', 'Обратное'].groupby(df.index.floor('D')).transform('sum') < 0.25 *
             df['Общая интенсивность автомобилей', 'Прямое'].groupby(df.index.floor('D')).transform('sum')) |
            (df.sum(axis=1) == 0) |
            (sum_non_passenger_cars['Сумма не-легковых', 'Обратное'] > 10 * df_tmp_back),
            'Данные НЕкорректны', 'Данные корректны')

        checker_data['Величина ошибки', 'Итого'] = np.where(
            (checker_data['Величина ошибки', 'Прямое'] == 'Данные НЕкорректны') |
            (checker_data['Величина ошибки', 'Обратное'] == 'Данные НЕкорректны'),
            'Данные НЕкорректны', 'Данные корректны')

        # Заведомо некорректные или отсуствующие данные: прямое направление
        straight_errors = sum(np.where(checker_data['Величина ошибки', 'Прямое'] == 'Данные НЕкорректны', 1, 0))
        # Заведомо некорректные или отсуствующие данные: обратное направление
        reverse_errors = sum(np.where(checker_data['Величина ошибки', 'Обратное'] == 'Данные НЕкорректны', 1, 0))

        # errors_statistics = pd.DataFrame({('Дата', 'Дата'): {},
        #                                   ('Величина ошибки', 'Логические'): {},
        #                                   ('Величина ошибки', 'Лишние данные'): {},
        #                                   ('Величина ошибки', 'Количество мотоциклов'): {},
        #                                   ('Величина ошибки', 'Некорректные данные Прямое'): {},
        #                                   ('Величина ошибки', 'Некорректные данные Обратное'): {}})
        self.errors_statistics['Дата', 'Дата'] = df.index
        self.errors_statistics.set_index(('Дата', 'Дата'), inplace=True)
        self.errors_statistics['Величина ошибки', 'Логические'] = baz
        self.errors_statistics['Величина ошибки', 'Лишние данные'] = foo
        if which_sample in ['Rosautodor_1', 'Rosautodor_2']:
            self.errors_statistics['Величина ошибки', 'Количество мотоциклов'] = bar
        self.errors_statistics['Величина ошибки', 'Некорректные данные Прямое'] = straight_errors
        self.errors_statistics['Величина ошибки', 'Некорректные данные Обратное'] = reverse_errors

        self.editor.insert(tk.END,
                           f"Некорректные данные: прямое направление: {straight_errors}\n"
                           f"Некорректные данные: обратное направление: {reverse_errors}\n")

        print(f'Количество логических ошибок в данных: {baz}')
        print(f'Величина ошибки между суммарным потоком и суммой по группам: {foo}%')
        if which_sample in ['Rosautodor_1', 'Rosautodor_2']:
            print(f'Кол-во зарегистрированных мотоциклов в исходных данных: {bar}')
        print(f'Некорректные данные: прямое направление: {straight_errors}')
        print(f'Некорректные данные: обратное направление: {reverse_errors}')
        # checker_data.reset_index(inplace=True)
        if which_sample in ['Rosautodor_1', 'Rosautodor_2']:
            logger_FDA.info(f'Количество логических ошибок в данных: {baz}\n'
                            f'\tВеличина ошибки между суммарным потоком и суммой по группам: {foo}%\n'
                            f'\tКол-во зарегистрированных мотоциклов в исходных данных: {bar}\n'
                            f'\tНекорректные данные: прямое направление: {straight_errors}\n'
                            f'\tНекорректные данные: обратное направление: {reverse_errors}')
        elif which_sample in ['Autodor_1', 'Autodor_2', 'Autodor_3']:
            logger_GK.info(f'Количество логических ошибок в данных: {baz}\n'
                           f'\tВеличина ошибки между суммарным потоком и суммой по группам: {foo}%\n'
                           f'\tНекорректные данные: прямое направление: {straight_errors}\n'
                           f'\tНекорректные данные: обратное направление: {reverse_errors}')
        return checker_data, self.errors_statistics

    def __find_missing_intervals_with_indices(self, df_in):
        """
        Функция для нахождения длины интервалов с пропущенными значениями.
        :param df_in:
        :return:
        """
        missing_intervals = []
        indices = []
        current_interval_length = 0
        start_index = None

        for i, value in enumerate(df_in['Количество']):
            if pd.isna(value):
                if current_interval_length == 0:
                    start_index = df_in.index[i]  # Запоминаем начальный индекс интервала
                current_interval_length += 1  # Увеличиваем длину интервала пропусков
            else:
                if current_interval_length > 0:
                    missing_intervals.append(current_interval_length)  # Сохраняем длину интервала
                    indices.append((start_index, df_in.index[i - 1]))  # Сохраняем индексы начала и конца интервала
                    current_interval_length = 0  # Сбрасываем длину интервала

        # Проверка на случай, если последний интервал пропусков в конце
        if current_interval_length > 0:
            missing_intervals.append(current_interval_length)
            indices.append((start_index, df_in.index[-1]))  # Сохраняем индексы для последнего интервала

        return missing_intervals, indices

    def __filling_gaps(self, df_in, date_time_idx, max_depth=10):
        """
        Функция для заполнения пропущенных значений с помощью простого среднего.
        Относительно NaN ищется значение неделю вперед/назад и от них берется среднее.
        Недостаток: ряд может иметь переменную сезонность (новогодние праздники, закрытие дороги, ремонт и т.п.).
        Из-за этого простое среднее может быть не совсем корректно.
        Есть еще всякие алгоритмы типа Калмана, MSTL, STL и даже простая интерполяция...
        (Далее также пробую декомпозицию временного ряда (MSTL) и скользящее среднее (window).)
        :param df_in: срез данных по направлению и типу автомобиля.
        :param date_time_idx: текущий индекс даты, где есть NaN
        :param max_depth: глубина поиска < 10 итераций.
        :return: найденное значение (среднее или np.nan, если значение не нашлось)
        """
        # думаю надо ли проставлять NaN'ы в начале года,
        # так как там сильная непредсказуемая динамика - исключаю из рассчетов
        df_copy = df_in.copy()
        # df_copy.loc['2024-01-01 00:59:59':'2024-01-08 23:59:59'] = np.nan

        found_before, found_after = None, None
        depth = 0
        current_idx = date_time_idx
        # df = df_total_long.loc[cond, 'Количество']
        while depth < max_depth:
            # Ищем неделю назад
            week_before = current_idx - pd.Timedelta(weeks=1)
            if week_before in df_copy.index and found_before is None:
                # изначально было:
                found_before = df_in.at[week_before] if not np.isnan(df_in.at[week_before]) else None
                # found_before = np.mean(df_copy.loc[week_before - pd.Timedelta(hours=1):
                #                               week_before + pd.Timedelta(hours=1)]) if not np.isnan(df_copy.at[week_before]) else None
            elif week_before not in df_copy.index:
                return np.nan
            # Ищем неделю вперёд
            week_after = current_idx + pd.Timedelta(weeks=1)
            if week_after in df_copy.index and found_after is None:
                # изначально было:
                found_after = df_in.at[week_after] if not np.isnan(df_in.at[week_after]) else None
                # found_after = np.mean(df_copy.loc[week_after - pd.Timedelta(hours=1)
                #                              :week_after + pd.Timedelta(hours=1)]) if not np.isnan(df_copy.at[week_after]) else None
            elif week_after not in df_copy.index:
                return np.nan
            # Логирование для отладки

            # Если оба значения найдены, возвращаем их среднее
            if found_before is not None and found_after is not None:
                # print(
                #     f"Дата : {current_idx},\nweek_before : {week_before}, found_before : {found_before},\nweek_after : {week_after}, found_after : {found_after}\n\n")
                return (found_before + found_after) / 2

            if np.isnan(df_copy.at[current_idx]):
                # Если не найдено ни одно значение, двигаемся дальше на одну неделю
                if found_before is None:
                    current_idx -= pd.Timedelta(weeks=1)  # Сдвигаемся на неделю назад
                elif found_after is None:
                    current_idx += pd.Timedelta(weeks=1)  # Сдвигаемся на неделю вперед
            else:
                # Если значение найдено, выходим из текущего прохода цикла
                break

            # Увеличиваем глубину
            depth += 1
        # попытка через декомпозицию временного ряда на тренд, сезонную и остаточную составляющие
        # for i in tqdm(types_vehicle):
        #     for j in directions:
        #         tmp_df = df_total_long.query(f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)
        #         res = MSTL(tmp_df.loc[:, 'Количество'].interpolate(method="linear"), periods=(24, 24*7)).fit()
        #         plt.rc("figure", figsize=(10, 10))
        #         plt.rc("font", size=5)
        #         res.plot()
        #
        #         seasonal_component = res.seasonal
        #         seasonal_component.head()
        #
        #         df_deseasonalised = tmp_df.loc[:, 'Количество'] - seasonal_component['seasonal_24'] - seasonal_component['seasonal_168']
        #         df_deseasonalised_imputed = df_deseasonalised.interpolate(method="linear")
        #         df_imputed = df_deseasonalised_imputed + seasonal_component['seasonal_24'] + seasonal_component['seasonal_168']
        #         df_imputed = df_imputed.to_frame().rename(columns={0: "Количество"})
        #         #         ax = df_imputed.plot(linestyle="-", marker=".", figsize=[10, 5], legend=None)
        #         #         ax = df_imputed[tmp_df.loc[:, 'Количество'].isnull()].plot(ax=ax, legend=None, marker=".", color="r")

        # способ через скользящее среднее (rolling). но не учитывается сезонность, поэтому использую другой алгоритм
        # for k in tqdm(types_vehicle):
        #     for l in directions:
        #         for i, j in zip(miss_interval, idx):
        #             tmp_df = df_total_long.query(f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)
        #             i=100
        #             j=(pd.Timestamp('2024-01-11 12:59:59'), pd.Timestamp('2024-01-15 15:59:59'))
        #             start, end = j
        #             print(i, j)
        #             print(start, end)
        #             tmp_df = tmp_df.assign(**{f'RollingMean_{i}': tmp_df.loc[start:end, 'Количество'].fillna(
        #                 tmp_df.loc[:, 'Количество'].shift(int(-i-1)).rolling(window='1h', closed='both', min_periods=1).mean())})
        #
        #             tmp_df = tmp_df.assign(**{f'RollingMedian_{i}': tmp_df.loc[start:end, 'Количество'].fillna(
        #                 tmp_df.loc[:, 'Количество'].shift(int(-i)).rolling(window='2h', closed='both', min_periods=1).median())})  # imputing using the median

        # Если не удалось найти значения, возвращаем NaN
        return np.nan

    # удаляю выбросы с помощью z-оценки
    def __zscore(self, s, window, sigma=3, return_all=False):
        # s = heh.replace(0, np.nan)['Количество']
        # s = s.to_frame()
        # window = '24h'
        roll = s.rolling(window=window, min_periods=1, center=True)

        avg = roll.mean()
        std = roll.std(ddof=0)
        z = s.sub(avg).div(std)
        m = z.between(-sigma, sigma)

        if return_all:
            return z, avg, std, m
        return s.where(m, np.nan)

    def fill_gaps_and_remove_outliers(self, df_total_long, detector_id, which_sample):
        df = df_total_long.copy()
        # зануляю те данные, которые заведомо НЕкорректны. Если "Прямое" - "Данные НЕкорректны", зануляю "Прямое",
        # если "Обратное" - зануляю "Обратное". Если "Итого" - зануляю "Итого"
        df['Количество'] = np.where(df['Корректность'] == 'Данные НЕкорректны',
                                    0, df['Количество'])

        types_vehicle = list(df['type_vehicle'].drop_duplicates())
        directions = list(df['direction'].drop_duplicates())
        outliers = []
        self.editor.insert(tk.END, f'Провожу обработку выбросов во временном ряду...\n')

        if which_sample in ['Rosautodor_1', 'Rosautodor_2']:
            logger_FDA.info(f'Провожу обработку выбросов во временном ряду...')
        elif which_sample in ['Autodor_1', 'Autodor_2', 'Autodor_3']:
            logger_GK.info(f'Провожу обработку выбросов во временном ряду...')

        print('Провожу обработку выбросов во временном ряду...')
        for i in tqdm(types_vehicle):
            for j in directions:
                tmp_df = df.query(f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)
                if np.isnan(tmp_df.loc[:, 'Количество']).all():
                    tmp_df.loc[:, "Количество"] = 0
                    outliers.append(tmp_df)
                    continue
                tmp_df.loc[:, 'Количество'] = self.__zscore(tmp_df['Количество'], '24h')
                outliers.append(tmp_df)

        outliers_long = pd.concat(outliers)

        imputed = []
        self.parent.update_idletasks()
        time.sleep(2)
        self.editor.insert(tk.END, f'Провожу дополнение временного ряда, заменяю пропуски...\n')

        if which_sample in ['Rosautodor_1', 'Rosautodor_2']:
            logger_FDA.info(f'Провожу дополнение временного ряда, заменяю пропуски...')
        elif which_sample in ['Autodor_1', 'Autodor_2', 'Autodor_3']:
            logger_GK.info(f'Провожу дополнение временного ряда, заменяю пропуски...')

        print('Провожу дополнение временного ряда, заменяю пропуски...')
        for i in tqdm(types_vehicle):
            for j in directions:
                # i = 'Общая интенсивность автомобилей'
                # j = 'Прямое'
                tmp_df = outliers_long.query(f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)

                miss_interval, idx = self.__find_missing_intervals_with_indices(
                    tmp_df.loc[tmp_df.index.isin(self.time_interval_cond)])

                # логика: +/- один час от интервала с пропусками заведомо считается плохим и поэтому также обнуляется
                for k, l in zip(miss_interval, idx):
                    start, end = l
                    # print(i, j)
                    # print(start, end)
                    first_index = start - dt.timedelta(hours=1)
                    last_index = end + dt.timedelta(hours=1)

                    if first_index in tmp_df.index:
                        tmp_df.loc[first_index, 'Количество'] = np.nan

                    if last_index in tmp_df.index:
                        tmp_df.loc[last_index, 'Количество'] = np.nan

                if np.isnan(tmp_df.loc[:, 'Количество']).all():
                    tmp_df.loc[:, "Количество"] = 0
                    imputed.append(tmp_df)
                    continue
                # алгоритм заполнения пропусков запускается через лямбду для каждой строки (.map) - эффективность так себе, но верно служит /это не много, но это хорошая работа/
                tmp_df.loc[:, "Количество"] = tmp_df.loc[:, 'Количество'].index.map(
                    lambda dt: self.__filling_gaps(tmp_df.loc[:, 'Количество'], dt) if np.isnan(
                        tmp_df.at[dt, 'Количество']) else tmp_df.at[
                        dt, 'Количество'])  # else df_total_long.loc[cond].at[dt, 'Количество'])

                imputed.append(tmp_df)

        imputed_df = pd.concat(imputed)  # датафрейм со всеми дополненными данными (выбросы и пропуски)
        bs_SSID, bs_intensivnosti = self.__calculate_statistics(imputed_df, detector_id, which_sample)

        # создаю датафрейм только по показателю общая интенсивность (для графиков)
        df_main_clear = []
        for j in directions:
            for i in imputed:
                tmp_df = i.query(f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'")
                if not tmp_df[(tmp_df.type_vehicle == 'Общая интенсивность автомобилей')
                              & (tmp_df.direction == j)].empty:
                    df_main_clear.append(i)
        df_main_clear = pd.concat(df_main_clear)

        return df_main_clear, bs_SSID, bs_intensivnosti

    def __calculate_statistics(self, imputed_df, detector_id, which_sample):
        """
        "В сечении" означает, что при обработке только по одному из направлений (прямому или обратному)
        ССИД принудительно увеличивается вдвое ("учитывается второе направление")

        :param imputed_df:
        :param which_sample:
        :return:
        """
        if which_sample in ['Rosautodor_1', 'Rosautodor_2']:
            logger_FDA.info(f'Считаю статистику...')
        elif which_sample in ['Autodor_1', 'Autodor_2', 'Autodor_3']:
            logger_GK.info(f'Считаю статистику...')

        if which_sample == 'Rosautodor_1':
            self.coef_sample = pd.read_excel(f'../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_fda_1')
        elif which_sample == 'Rosautodor_2':
            self.coef_sample = pd.read_excel(f'../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_fda_2')
        elif which_sample == 'Autodor_1':
            self.coef_sample = pd.read_excel(f'../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_gk_1')
        elif which_sample == 'Autodor_2':
            self.coef_sample = pd.read_excel(f'../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_gk_2')
        elif which_sample == 'Autodor_3':
            self.coef_sample = pd.read_excel(f'../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_gk_3')

        # Кол-во ТС за полный год (для оценки ССИД) - сразу в сечении
        amount_per_year_vehicle = imputed_df.loc[self.time_interval_cond].groupby(by=['direction', 'type_vehicle'])[
            'Количество'].sum().unstack(0)
        amount_per_year_vehicle.loc[:, ['Обратное', 'Прямое']] *= 2
        amount_per_year_vehicle = amount_per_year_vehicle.T.stack()

        imputed_df['y_m_d'] = imputed_df.index.strftime('%Y-%m-%d')
        imputed_df['y_m_d'] = pd.to_datetime(imputed_df['y_m_d'])
        # imputed_df.dtypes

        # Считаю статистику по СУТОЧНОЙ интенсивности
        imputed_df_day = imputed_df.loc[self.time_interval_cond].groupby(by=['y_m_d', 'direction'])['Количество'].sum().unstack(1)

        # Максимум за сутки
        max_value = np.max(imputed_df_day, axis=0).replace(0, np.nan)
        max_value.name = 'максимум за сутки'
        # Минимум за сутки
        min_value = np.min(imputed_df_day, axis=0)
        min_value.name = 'минимум за сутки'
        # Среднее значение за сутки
        mean_value = np.mean(imputed_df_day, axis=0)
        mean_value.name = 'среднее за сутки'
        # Коэффициент перехода (максимум делить на среднее) - куда переход? - непонятки
        coeff_trans = max_value / mean_value
        coeff_trans.name = 'коэф прехода (макс'

        # ПОДУМОТЬ Среднее суточное значение, физ.ед. (в сечении) - что если естть данные только для одного направления
        # mean_cross_section = np.where(mean_value['Прямое'] == 0,
        #                               np.where(mean_value.index == 'Итого', mean_value['Итого'], mean_value),
        #                               np.where(mean_value['Обратное'] == 0,
        #                                        np.where(mean_value.index == 'Итого', mean_value['Итого'], mean_value),
        #                                        np.where(mean_value.index == 'Итого', mean_value['Итого'], mean_value * 2)))

        # Среднее суточное значение, физ.ед.
        mean_cross_section = pd.Series(np.where(mean_value.index == 'Итого', mean_value['Итого'], mean_value * 2),
                                       index=mean_value.index)
        mean_cross_section.name = 'среднесуточное в сечении'

        # ССИД в летний период
        mean_summer = np.mean(imputed_df_day['2024-06-01':'2024-08-31'], axis=0)
        mean_cross_section_summer = pd.Series(np.where(mean_summer.index == 'Итого', mean_summer['Итого'], mean_summer * 2),
                                              index=mean_summer.index)
        mean_cross_section_summer.name = 'ССИД летом'

        # ССИД в зимний период
        mean_winter = np.mean(pd.concat([imputed_df_day['2024-01-01':'2024-02-29'],
                                         imputed_df_day['2024-12-01':'2024-12-31']], axis=0), axis=0)
        mean_cross_section_winter = pd.Series(np.where(mean_winter.index == 'Итого', mean_winter['Итого'], mean_winter * 2),
                                              index=mean_winter.index)
        mean_cross_section_winter.name = 'ССИД зимой'

        # ССИД в межсезонье
        mean_demiseasons = np.mean(pd.concat([imputed_df_day['2024-03-01':'2024-05-31'],
                                         imputed_df_day['2024-09-01':'2024-11-30']], axis=0), axis=0)
        mean_cross_section_demiseasons = pd.Series(np.where(mean_demiseasons.index == 'Итого', mean_demiseasons['Итого'], mean_demiseasons * 2),
                                                   index=mean_demiseasons.index)
        mean_cross_section_demiseasons.name = 'ССИД межсезонье'

        # Максимальный суточный поток, физ.ед.
        max_cross_section = pd.Series(np.where(max_value.index == 'Итого', max_value['Итого'], max_value * 2),
                                      index=max_value.index)
        max_cross_section.name = 'макс суточный поток'
        # Минимальный суточный поток, физ.ед.
        min_cross_section = pd.Series(np.where(min_value.index == 'Итого', min_value['Итого'], min_value * 2),
                                      index=min_value.index)
        min_cross_section.name = 'мин суточный поток'

        imputed_df['hour'] = imputed_df.index.strftime("%H").astype(int)

        # Максимальная часовая интенсивность в дневной период, физ.ед.
        max_hour_day = imputed_df[imputed_df.index.isin(self.time_interval_cond) & imputed_df['hour'].between(7, 22)].groupby(
            by=['direction'])['Количество'].max()
        max_hour_day_cross_section = pd.Series(np.where(max_hour_day.index == 'Итого', max_hour_day['Итого'], max_hour_day * 2),
                                               index=max_hour_day.index)
        max_hour_day_cross_section.name = 'макс часов интенсивность днем'

        # Максимальная часовая интенсивность в ночной период, физ.ед.
        max_hour_night = imputed_df[imputed_df.index.isin(self.time_interval_cond) & imputed_df['hour'].isin([23, 0, 1, 2, 3, 4, 5, 6])].groupby(by=['direction'])['Количество'].max()
        max_hour_night_cross_section = pd.Series(np.where(max_hour_night.index == 'Итого', max_hour_night['Итого'], max_hour_night * 2),
                                                 index=max_hour_night.index)
        max_hour_night_cross_section.name = 'макс часов интенсивность ночью'
        # День, соответствующий суточному максимуму
        # создаю новый series для хранения результата (день, где был максимум)
        days_of_max = pd.Series(index=imputed_df_day.columns, dtype=object)

        for idx in imputed_df_day.columns:
            # idx = foo.columns[0]
            col = idx
            matches = imputed_df_day[col].isin([max_value[idx]])  # проверяю совпадения с соответствующим значением из max_value

            if matches.sum() == 1:  # сохраняю день
                days_of_max[idx] = pd.to_datetime(imputed_df_day[idx].index[matches].date).strftime('%Y-%m-%d').values
            elif matches.sum() >= 2:  # если совпадений два или более, сохраняю список дней
                days_of_max[idx] = [pd.to_datetime(imputed_df_day[idx].index[matches].date).strftime('%Y-%m-%d').values]
            else:
                days_of_max[idx] = pd.Timestamp('NaT').to_pydatetime()  # проставляю np.nan, если совпадений нет
        days_of_max.name = 'день максимума'

        imputed_df_TG = imputed_df.loc[self.time_interval_cond].reset_index().merge(self.coef_sample,
                                                                                 how='left',
                                                                                 on='type_vehicle').set_index('Дата')
        # пересчитываю количество автомобилей уже на тарифные группы
        imputed_df_TG['Количество_ТГ'] = np.prod([imputed_df_TG['Количество'], imputed_df_TG['coeff']], axis=0)

        imputed_df_TG['month'] = imputed_df_TG.index.strftime("%B")
        imputed_df_TG['day_of_week'] = imputed_df_TG.index.strftime("%A")
        imputed_df_TG['hour'] = imputed_df_TG.index.strftime("%H")

        # считаю коэффициенты неравномерности
        # по месяцам
        coeff_by_month = imputed_df_TG.groupby(by=['month', 'direction', 'TG'])['Количество_ТГ'].sum().unstack([1, 2])
        coeff_by_month = coeff_by_month.div(
            coeff_by_month.sum(axis=0),
            axis=1)
        # по дням недели
        coeff_by_weekday = imputed_df_TG.groupby(by=['day_of_week', 'direction', 'TG'])['Количество_ТГ'].sum().unstack(
            [1, 2])
        coeff_by_weekday = coeff_by_weekday.div(
            coeff_by_weekday.sum(axis=0),
            axis=1)
        # по часам
        coeff_by_hour = imputed_df_TG.groupby(by=['hour', 'direction', 'TG'])['Количество_ТГ'].sum().unstack([1, 2])
        coeff_by_hour = coeff_by_hour.div(
            coeff_by_hour.sum(axis=0),
            axis=1)

        # Величина среднегодовой суточной интенсивности В СЕЧЕНИИ по тарифным группам
        avg_annual_per_24_h_TG = amount_per_year_vehicle.unstack(0).reset_index().merge(self.coef_sample,
                                                                                 how='left',
                                                                                 on='type_vehicle').set_index('type_vehicle')
        avg_annual_per_24_h_TG.loc[:, 'Итого'] = avg_annual_per_24_h_TG.loc[:, 'Итого'] * avg_annual_per_24_h_TG.loc[:, 'coeff']
        avg_annual_per_24_h_TG.loc[:, 'Обратное'] = avg_annual_per_24_h_TG.loc[:, 'Обратное'] * avg_annual_per_24_h_TG.loc[:, 'coeff']
        avg_annual_per_24_h_TG.loc[:, 'Прямое'] = avg_annual_per_24_h_TG.loc[:, 'Прямое'] * avg_annual_per_24_h_TG.loc[:, 'coeff']
        if which_sample in ['Autodor_1', 'Autodor_2', 'Autodor_3']:
            if 'неопознаные тс' not in avg_annual_per_24_h_TG.index:
                avg_annual_per_24_h_TG = pd.concat([avg_annual_per_24_h_TG, pd.DataFrame({'Итого': 0.0,
                                                                             'Обратное': 0.0,
                                                                             'Прямое': 0.0,
                                                                             'TG': 'Неопознанные_ТГ',
                                                                             'coeff': 1.0}, index=['неопознаные тс'])])
                                                   # .set_index('неопознаные тс', axis=0)
                # avg_annual_per_24_h_TG = avg_annual_per_24_h_TG[avg_annual_per_24_h_TG['TG'] != 'Unknow ТГ']
            # logger_GK.info(f'Считаю статистику...')
        avg_annual_per_24_h_TG = avg_annual_per_24_h_TG.groupby('TG')[['Итого', 'Прямое', 'Обратное']].sum() / 365

        # avg_annual_per_24_h_TG = imputed_df_TG.groupby(by=['direction', 'TG'])['Количество_ТГ'].sum() / 365 # а можно было в одну строчку D;    ;-----;

        # global bs_SSID
        # global bs_intensivnosti
        # SSID - транслитерация от СреднеСуточная Интенсивность Движения
        self.basic_stats_SSID = pd.concat([avg_annual_per_24_h_TG.astype(str).replace('nan', '0.0'),
                                      max_value.to_frame().T.astype(str).replace('nan', '0.0'),
                                      min_value.to_frame().T.astype(str).replace('nan', '0.0'),
                                      mean_value.to_frame().T.astype(str).replace('nan', '0.0'),
                                      coeff_trans.to_frame().T.astype(str).replace('nan', '0.0'),
                                      mean_cross_section.to_frame().T.astype(str).replace('nan', '0.0'),
                                      mean_cross_section_summer.to_frame().T.astype(str).replace('nan', '0.0'),
                                      mean_cross_section_winter.to_frame().T.astype(str).replace('nan', '0.0'),
                                      mean_cross_section_demiseasons.to_frame().T.astype(str).replace('nan', '0.0'),
                                      max_cross_section.to_frame().T.astype(str).replace('nan', '0.0'),
                                      min_cross_section.to_frame().T.astype(str).replace('nan', '0.0'),
                                      max_hour_day_cross_section.to_frame().T.astype(str).replace('nan', '0.0'),
                                      max_hour_night_cross_section.to_frame().T.astype(str).replace('nan', '0.0'),
                                      days_of_max.to_frame().T.replace('NaT', '0.0')], axis=0).reset_index()
        self.basic_stats_SSID.index = [detector_id] * len(self.basic_stats_SSID)

        self.basic_stats_intensivnosti = pd.concat([coeff_by_month,
                                               coeff_by_weekday,
                                               coeff_by_hour], axis=0).reset_index()
        self.basic_stats_intensivnosti.index = [detector_id] * len(self.basic_stats_intensivnosti)

        return self.basic_stats_SSID, self.basic_stats_intensivnosti

    def plot_graphs(self, company, df_main_clear, df_total_long, cur_year, file, freq='d'):
        df_raw = df_total_long.copy()
        df_clear = df_main_clear.copy()

        self.parent.update_idletasks()
        time.sleep(2)
        self.editor.insert(tk.END, f'Рисую графики для файла {file}...\n')
        if company == 'ФДА':
            logger_FDA.info(f'Рисую графики для файла {file}...')
        elif company == 'ГК':
            logger_GK.info(f'Рисую графики для файла {file}...')

        # fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(20, 10))
        fig = plt.figure(figsize=(20, 10))
        fig.suptitle(f'{file}', fontsize=13)
        fig.subplots_adjust(hspace=0.5)  # увеличивает расстояния между графиками (сабплотсами внутри fig) по вертикали
        gs = GridSpec(3, 4, figure=fig)  # три строки (три рисунка по вертикали) и два столбца (два рисунка по горизонтали)
        ax1 = fig.add_subplot(gs[0, :-1])
        ax2 = fig.add_subplot(gs[1, :-1])
        ax3 = fig.add_subplot(gs[-1, 0:2])  # -1 это последняя строка (по-другому можно записать 2)
        ax4 = fig.add_subplot(gs[-1, 2:4])  # -1 это последняя строка (по-другому можно записать 2)
        ax5 = fig.add_subplot(gs[0:2, -1])
        ax5.set_axis_off()
        # ax5.text(0.5, 0.5, f'my text')

        colors_line = ['cornflowerblue', 'darkblue']
        colors_dots = ['lime', 'limegreen']
        directions = ['Прямое', 'Обратное']

        # Рисую исходные данные
        for idx, j in enumerate(directions):
            tmp_df_raw = df_raw.query(
                f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'").replace(np.nan, 0).copy()
            tmp_df_raw.loc[:, 'y_m_d'] = tmp_df_raw.index.strftime('%Y-%m-%d').values
            tmp_df_raw.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df_raw['y_m_d'])

            tmp_df_raw_day = tmp_df_raw.loc[self.time_interval_cond].groupby(by=['y_m_d', 'direction'])[
                'Количество'].sum().unstack(1)

            if not tmp_df_raw.empty:
                if freq == 'd':
                    # tmp_df_raw_day.plot(style=colors_line[idx % 2], ax=ax2, legend=False, label=f'{j} направление')
                    ax1.plot(pd.to_datetime(tmp_df_raw_day.index),
                             tmp_df_raw_day.astype(float),
                             color=colors_line[idx % 2],
                             label=f'{j} направление',
                             linewidth=0.5, zorder=1)
                    ax1.xaxis.set_major_locator(MonthLocator())
                    ax1.xaxis.set_major_formatter(DateFormatter('%Y-%m'))
                    ax1.autoscale()
                    ax1.set_ylim(0, None)

                elif freq == 'h':
                    ax1.plot(pd.to_datetime(tmp_df_raw.loc[tmp_df_raw.index.isin(self.time_interval_cond), 'Количество'].index),
                             tmp_df_raw.loc[tmp_df_raw.index.isin(self.time_interval_cond), 'Количество'].astype(float),
                             color=colors_line[idx % 2],
                             label=f'{j} направление',
                             linewidth=0.5, zorder=1)
                    ax1.xaxis.set_major_locator(MonthLocator())
                    ax1.xaxis.set_major_formatter(DateFormatter('%Y-%m'))
                    ax1.autoscale()
                    ax1.set_ylim(0, None)

                ax1.set_title('Суточная интенсивность: исходные данные')  # Заголовок второго графика
                ax1.legend(loc='best')  # Указываем, где разместить легенду
                ax1.set_ylabel('Количество')  # Подпись оси Y
                ax1.set_xlabel('Время')  # Подпись оси X

        # Рисую дополненные данные
        for idx, j in enumerate(directions):
            # j = 'Прямое'
            tmp_df = df_clear.query(f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'").copy()
            tmp_df.loc[:, 'y_m_d'] = tmp_df.index.strftime('%Y-%m-%d').values
            tmp_df.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df['y_m_d'])
            tmp_df_day = tmp_df.loc[self.time_interval_cond].groupby(by=['y_m_d', 'direction'])[
                'Количество'].sum().unstack(1)

            # tmp_df_raw - этот нужен для отрисовки пропусков в виде маркеров на графике зеленым цветом (где nan, там рисуем метку)
            tmp_df_raw = df_raw.query(
                f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'").replace(0, np.nan).copy()
            tmp_df_raw.loc[:, 'y_m_d'] = tmp_df_raw.index.strftime('%Y-%m-%d').values
            tmp_df_raw.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df_raw['y_m_d'])

            # для зеленых точек - показывают, где был пропуск
            tmp_df_raw_day = tmp_df_raw.loc[self.time_interval_cond].groupby(by=['y_m_d', 'direction'])[
                'Количество'].sum().unstack(1).replace(0, np.nan)

            if not tmp_df.empty:
                if freq == 'd':
                    # tmp_df_day.plot(style=colors_line[idx % 2], ax=ax1, legend=False, label=f'{j} направление')

                    # tmp_df_day[tmp_df_raw_day.isna().any(axis=1)].reindex(tmp_df_day.index).replace(0, np.nan) \
                    #     .plot(style='o', color=colors_dots[idx % 2], ax=ax1, legend=False, label=f'{j} направление')

                    ax2.plot(pd.to_datetime(tmp_df_day.index), tmp_df_day.replace(np.nan, 0).astype(float),
                             color=colors_line[idx % 2],
                             label=f'{j} направление',
                             linewidth=0.5, zorder=1)

                    # ax2.scatter(pd.to_datetime(tmp_df_day[tmp_df_raw_day.isna().any(axis=1)].reindex(tmp_df_day.index).index),
                    #             tmp_df_day[tmp_df_raw_day.isna().any(axis=1)].reindex(tmp_df_day.index).replace(0, np.nan).astype(float),
                    #             color=colors_dots[idx % 2], marker='o', s=10, label=f'{j} направление', zorder=2)
                    ax2.autoscale()
                    ax2.set_ylim(0, None)
                    ax2.xaxis.set_major_locator(MonthLocator())
                    ax2.xaxis.set_major_formatter(DateFormatter('%Y-%m'))

                elif freq == 'h':
                    # ax2.scatter(pd.to_datetime(tmp_df[tmp_df.index.isin(self.time_interval_cond) & tmp_df_raw.isna().any(axis=1)].loc[:,
                    #             'Количество'].index),
                    #             tmp_df[tmp_df.index.isin(self.time_interval_cond) & tmp_df_raw.isna().any(axis=1)].loc[:,
                    #             'Количество'].replace(0, np.nan).astype(float),
                    #             color=colors_dots[idx % 2], marker='o', s=10, label=f'{j} направление', zorder=2)

                    ax2.plot(pd.to_datetime(tmp_df.loc[tmp_df.index.isin(self.time_interval_cond), 'Количество'].index),
                             tmp_df.loc[tmp_df.index.isin(self.time_interval_cond), 'Количество'].astype(float),
                             color=colors_line[idx % 2],
                             label=f'{j} направление',
                             linewidth=0.5, zorder=1)
                    ax2.autoscale()
                    ax2.set_ylim(0, None)
                    ax2.xaxis.set_major_locator(MonthLocator())
                    ax2.xaxis.set_major_formatter(DateFormatter('%Y-%m'))

                ax2.set_title(
                    'Суточная интенсивность: скорректированные (дополненные) данные')  # Заголовок первого графика
                ax2.legend(loc='best')  # Указываем, где разместить легенду
                ax2.set_ylabel('Дополненное количество')  # Подпись оси Y
                ax2.set_xlabel('Время')  # Подпись оси X

        # Рисую графики неравномерностей
        basic_stats_intensivnosti_long = self.basic_stats_intensivnosti.reset_index().set_index(['level_0', 'index']) \
            .melt(ignore_index=False).reset_index(names=['detector_id', 'type'])

        for idx, j in enumerate(['Итого']):# enumerate(directions):  # неравномерности только по сумме направлений
            # создаю графики для недельной и часовой интенсивности
            # j = 'Прямое'
            weekdays_list = ["Monday", "Tuesday", "Wednesday", "Thursday",
                             "Friday", "Saturday", "Sunday"]
            hours_list = ['00', '01', '02', '03', '04', '05', '06', '07', '08', '09', '10',
                          '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21',
                          '22', '23']
            colors_intens = {'Прямое': ['red', 'salmon', 'coral', 'tomato'],
                             # ['red', 'blue', 'orange', 'green', 'purple'],
                             'Обратное': ['blue', 'skyblue', 'royalblue', 'dodgerblue'],
                             # ['teal', 'deeppink', 'cyan', 'yellow', 'peru'],
                             'Итого': ['darkblue', 'red', 'orange', 'green']}  # ['orange', 'goldenrod', 'darkorange', 'gold', 'khaki']}

            for i, tg_name in enumerate(['ТГ-1', 'ТГ-2', 'ТГ-3', 'ТГ-4']):
                # tg_name = 'Все ТГ'
                weekday = basic_stats_intensivnosti_long.query(
                    f"direction == '{j}' and type in {weekdays_list} and TG == '{tg_name}'") \
                    .set_index('type').reindex(weekdays_list).reset_index()  # .replace(np.nan, 0)
                hour = basic_stats_intensivnosti_long.query(
                    f"direction == '{j}' and type in {hours_list} and TG == '{tg_name}'")
                # for k in range(15):
                #     color_index = k // 5
                if not weekday.empty:
                    ax3.plot(hour['type'].astype(str),
                             hour['value'].astype(float),
                             color=colors_intens[j][i % 5],
                             label=f'{tg_name}, {j}')
                    ax3.autoscale()
                    ax3.set_ylim(0, None)

                    ax4.plot(weekday['type'].astype(str),
                             weekday['value'].astype(float),
                             color=colors_intens[j][i % 5],
                             label=f'{tg_name}, {j}')
                    # ax4.autoscale()
                    # ax4.set_ylim(0, None)
        ax3.set_title('Суточная неравномерность')
        ax3.yaxis.set_major_formatter(mtick.PercentFormatter(1, decimals=0))
        ax3.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=5, fontsize=9)
        ax3.set_ylabel('Процент')  # Подпись оси Y
        ax3.set_xlabel('Часы')  # Подпись оси X

        ax4.set_title('Неравномерность по дням недели')
        ax4.yaxis.set_major_formatter(mtick.PercentFormatter(1, decimals=0))
        ax4.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=5, fontsize=9)
        ax4.set_ylabel('Процент')  # Подпись оси Y
        ax4.set_xlabel('День недели')  # Подпись оси X

        ax5.axis([0, 10, 0, 10])
        ax5.text(1, 5, f"Максимальное количество замеров: {self.small_statistics[0]}\n"
                       f"Фактическое количество замеров: {self.small_statistics[1]}\n"
                       f'Полнота данных: {self.small_statistics[2]:,.1%}\n\n'
                       f"Логические ошибки в данных: {self.errors_statistics['Величина ошибки', 'Логические'].iloc[0]}\n"
                       f"Лишние данные: {self.errors_statistics['Величина ошибки', 'Лишние данные'].iloc[0]}\n"
                       f"Некорректные ПРЯМОЕ: {self.errors_statistics['Величина ошибки', 'Некорректные данные Прямое'].iloc[0]}\n"
                       f"Некорректные ОБРАТНОЕ:  {self.errors_statistics['Величина ошибки', 'Некорректные данные Обратное'].iloc[0]}",
                 fontsize=15)

        plt.rcParams['font.size'] = 13
        # plt.legend(loc='best')
        # plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(f'../raw_data/{company}/{cur_year}/Графики/' + 'PIC' + file[3:len(file) - 5] + '.png')
        # plt.show()
