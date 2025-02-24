import datetime as dt
import os
import time
import re

import numpy as np
import pandas as pd
import xlwings as xw
import openpyxl
from statsmodels.tsa.seasonal import MSTL
import matplotlib

matplotlib.use('TkAgg')
from matplotlib import pyplot as plt
from matplotlib.ticker import MaxNLocator
from matplotlib.dates import MonthLocator, DateFormatter
from matplotlib import colors as mcolors
# plt.switch_backend('agg')
from tqdm import tqdm
import tkinter as tk
from tkinter import *
from tkinter import messagebox

import detectors


class Checking:

    def __init__(self, editor, parent):
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
        # self.cur_year = 2024
        # self.file = ''  # 'PRE_км 198+500 а.д М-10 Россия Москва – Санкт-Петербург .xlsx'
        # self.directory_pre = f'../raw_data/{str(cur_year)}/Первичная обработка/'
        # self.df = self.open_and_read_file()

    def open_and_read_file(self, directory_pre, file, which_sample):
        try:
            self.wb_data = xw.Book(os.path.join(directory_pre, file))
            wb = openpyxl.load_workbook(os.path.join(directory_pre, file))
            sheet = wb.active
            last_row = len(sheet['A'])
            if which_sample == 'Rosautodor_1':
                my_values = self.wb_data.sheets[0].range('A7:Y' + str(last_row)).options(ndim=2).value
            elif which_sample == 'Rosautodor_2':
                my_values = self.wb_data.sheets[0].range('A7:AB' + str(last_row)).options(ndim=2).value

            df = pd.DataFrame(my_values)
            df.fillna(0, inplace=True)
            return df
        except Exception as e:
            print(f"Ошибка при чтении файла {file}: {e}")
            tk.messagebox.showerror(title="ALERT",
                                    message=f"Ошибка при чтении файла {file}: {e}")
            return None

    def make_long(self, df, cur_year, file, which_sample):
        self.editor.insert(tk.END, f'Превращаю данные файла {file} в длинный формат...\n')
        self.editor.see(END)
        # if list(df.tail(3)[0])[0] == 'Итого':
            # df.iloc[:len(df) - (df[df.columns[0]].to_numpy() == 'Итого')[::-1].argmax()]
        df = df.iloc[:df[df[df.columns[0]] == 'Итого'].index[0]].copy()
            # df.drop(df.tail(3).index, inplace=True)  # удаляю последние три строки ['Итого', 'Среднее', '%']
        # elif list(df.tail(3)[0])[0] != 'Итого' and list(df.tail(3)[0])[1] == 'Среднее':
        #     df = df.iloc[:df[df[df.columns[0]] == 'Среднее'].index[0]]
        #     # df.drop(df.tail(2).index, inplace=True)
        # elif list(df.tail(3)[0])[0] != 'Итого' and list(df.tail(3)[0])[1] != 'Среднее' and list(df.tail(3)[0])[
        #     2] == '%':
        #     df = df.iloc[:df[df[df.columns[0]] == '%'].index[0]]
        #     # df.drop(df.tail(1).index, inplace=True)
        # else:
        #     pass

        if which_sample == 'Rosautodor_1':
            self.column_names = self.wb_data.sheets['Исходные данные'].range('A3:Y4').value
        elif which_sample == 'Rosautodor_2':
            self.column_names = self.wb_data.sheets['Исходные данные'].range('A3:AB4').value

        # self.column_names = self.wb_data.sheets['Исходные данные'].range('A3:AB4').value
        self.column_names = pd.DataFrame(self.column_names).ffill(axis=1).ffill(axis=0)
        df.columns = pd.MultiIndex.from_arrays(self.column_names[:2].values, names=['type_vehicle', 'direction'])
        df[('Дата', 'Дата')] = pd.to_datetime(df[('Дата', 'Дата')], format='%d.%m.%Y %H:%M:%S')
        df.iloc[:, 1:] = df.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

        # принудительно делаем из object - float64
        for col in df.columns[1:]:
            if isinstance(df[col].dtype, object):
                df[col] = df[col].astype(float)

        tmp_df = self.__add_previous_december_data(cur_year, file, which_sample)
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

        max_measured = 8784
        fact_measured = sum(
            df[('Общая интенсивность автомобилей', 'Итого')][self.time_interval_cond].replace(0, np.nan).notna())
        fullness = fact_measured / max_measured
        self.editor.insert(tk.END, f'Максимальное количество замеров: {max_measured}\n'
                                   f'Фактическое количество замеров: {fact_measured}\n'
                                   f'Полнота данных: {fullness}\n')
        self.editor.see(END)

        checker_data, self.errors_statistics = self.__check_correct_data(df, which_sample)

        df_long = df.melt(ignore_index=False, value_name='Количество').reset_index().rename(
            columns={('Дата', 'Дата'): 'Дата', 'variable_0': 'type_vehicle', 'variable_1': 'direction'})
        checker_data_long = checker_data.melt(ignore_index=False, col_level=1, value_vars=['Обратное', 'Прямое'],
                                              var_name='direction', value_name='Корректность').reset_index().rename(
            columns={('Дата', 'Дата'): "Дата"})

        df_total_long = df_long.merge(checker_data_long, how='left', on=['Дата',
                                                                         'direction'])  # .rename(columns={'Значение_x': 'Количество', 'Значение_y': 'Корректность'})
        # тут надо повнимательнее! может быть переписать этот кусочек... потому что эти два показателя залезают
        # в данные только при первом варианте (загрузка и скорость)
        df_total_long = df_total_long[
            (df_total_long.type_vehicle != 'Загрузка, %') & (df_total_long.type_vehicle != 'Скорость, км/ч')]

        df_total_long.set_index('Дата', inplace=True)
        # df_total_long.
        return df_total_long

    def __add_previous_december_data(self, cur_year, file, which_sample):
        previous_year_files = os.listdir(f'../raw_data/{str(int(cur_year) - 1)}/Первичная обработка')
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
                    self.editor.see(END)
                    # tk.messagebox.showerror(title="ALERT",
                    #                         message=f"Файл за предыдущий год найден! {i}")
                    if which_sample == 'Rosautodor_1':
                        wb_prev_yr_file = pd.read_excel(
                            os.path.join(f'../raw_data/{str(int(cur_year) - 1)}/Первичная обработка/', file)).iloc[:,
                                          0:-4]
                    elif which_sample == 'Rosautodor_2':
                        wb_prev_yr_file = pd.read_excel(
                            os.path.join(f'../raw_data/{str(int(cur_year) - 1)}/Первичная обработка/', file)).iloc[:,
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
            self.editor.see(END)
            print(f"Не удалось найти файл {file} за прошлый год: {e}")
            tk.messagebox.showerror(title="ALERT",
                                    message=f"Не удалось найти файл {file} за прошлый год: {e}")
            return wb_prev_yr_file

    def __check_correct_data(self, df, which_sample):
        self.editor.insert(tk.END, f'Проверяю корректность данных и считаю статистику... \n')
        self.editor.see(END)
        sum_non_passenger_cars = pd.DataFrame({('Дата', 'Дата'): {},
                                               ('Сумма не-легковых', 'Итого'): {},
                                               ('Сумма не-легковых', 'Прямое'): {},
                                               ('Сумма не-легковых', 'Обратное'): {}})
        sum_non_passenger_cars['Дата', 'Дата'] = df.index
        sum_non_passenger_cars.set_index(('Дата', 'Дата'), inplace=True)
        try:
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
        except KeyError as e:
            print('Плашка')
            tk.messagebox.showerror(title="ALERT",
                                    message=f"Ошибка в наименованиях столбцов! {e}\n"
                                            f"Проверьте заголовки в данных. Возможно не хватает столбцов")
            return

        error_amount_direction = pd.DataFrame({('Дата', 'Дата'): {},
                                               ('Величина ошибки', 'Прямое'): {},
                                               ('Величина ошибки', 'Обратное'): {}})
        error_amount_direction['Дата', 'Дата'] = df.index
        error_amount_direction.set_index(('Дата', 'Дата'), inplace=True)

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
        logical_errors['Дата', 'Дата'] = df.index
        logical_errors.set_index(('Дата', 'Дата'), inplace=True)

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
                                                                    df['Легковые большие (4-6 м)', 'Прямое'] - \
                                                                    df['Малые груз. (6-9 м)', 'Обратное'] - \
                                                                    df['Грузовые (9-13 м)', 'Обратное'] - \
                                                                    df['Груз. большие (13-22 м)', 'Обратное'] - \
                                                                    df['Автопоезда (22-30 м)', 'Обратное'] - \
                                                                    df['Автобусы', 'Обратное'] - \
                                                                    df['Мотоциклы', 'Обратное']

        # Количество логических ошибок в данных
        baz = logical_errors.sum(axis=0).sum()
        # Величина ошибки между суммарным потоком и суммой по группам
        # ("лишние" данные по направлениям, деленное на Общую интенсивность авто по направлениям)
        foo = error_amount_direction.sum(axis=0).sum() / (df['Общая интенсивность автомобилей', 'Прямое'] +
                                                          df['Общая интенсивность автомобилей', 'Обратное']).sum()
        # Кол-во зарегистрированных мотоциклов в исходных данных
        bar = (df['Мотоциклы', 'Прямое'] + df['Мотоциклы', 'Обратное']).sum()

        checker_data = pd.DataFrame({('Дата', 'Дата'): {},
                                     ('Величина ошибки', 'Прямое'): {},
                                     ('Величина ошибки', 'Обратное'): {}})
        checker_data['Дата', 'Дата'] = df.index
        checker_data.set_index(('Дата', 'Дата'), inplace=True)
        checker_data['Величина ошибки', 'Прямое'] = np.where(
            (df['Общая интенсивность автомобилей', 'Прямое'] < 0.25 * df[
                'Общая интенсивность автомобилей', 'Обратное']) |
            (df.sum(axis=1) == 0) |  # .loc[(df.index > '2024-01-11 11:59:59') & (df.index < '15.01.2024 16:59:59')]
            (sum_non_passenger_cars['Сумма не-легковых', 'Прямое'] > 10 * (
                    df['Общая интенсивность автомобилей', 'Прямое'] + df[
                'Общая интенсивность автомобилей', 'Обратное'])), 'Данные НЕкорректны', 'Данные корректны')

        checker_data['Величина ошибки', 'Обратное'] = np.where(
            (df['Общая интенсивность автомобилей', 'Обратное'] < 0.25 * df[
                'Общая интенсивность автомобилей', 'Прямое']) |
            (df.sum(axis=1) == 0) |
            (sum_non_passenger_cars['Сумма не-легковых', 'Обратное'] > 10 * (
                    df['Общая интенсивность автомобилей', 'Прямое'] + df[
                'Общая интенсивность автомобилей', 'Обратное'])), 'Данные НЕкорректны', 'Данные корректны')

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
        self.errors_statistics['Величина ошибки', 'Количество мотоциклов'] = bar
        self.errors_statistics['Величина ошибки', 'Некорректные данные Прямое'] = straight_errors
        self.errors_statistics['Величина ошибки', 'Некорректные данные Обратное'] = reverse_errors

        self.editor.insert(tk.END,
                           f"Некорректные данные: прямое направление: {straight_errors}\n"
                           f"Некорректные данные: обратное направление: {reverse_errors}\n")

        print(f'Количество логических ошибок в данных: {baz}')
        print(f'Величина ошибки между суммарным потоком и суммой по группам: {foo}%')
        print(f'Кол-во зарегистрированных мотоциклов в исходных данных: {bar}')
        print(f'Некорректные данные: прямое направление: {straight_errors}')
        print(f'Некорректные данные: обратное направление: {reverse_errors}')
        # checker_data.reset_index(inplace=True)

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
        Из-за этого среднее может быть не совсем корректно.
        (Далее также пробую декомпозицию временного ряда и скользящее среднее.)
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
    def __zscore(self, s, window, thresh=3, return_all=False):
        # s = heh.replace(0, np.nan)['Количество']
        # s = s.to_frame()
        # window = '24h'
        roll = s.rolling(window=window, min_periods=1, center=True)

        avg = roll.mean()
        std = roll.std(ddof=0)
        z = s.sub(avg).div(std)
        m = z.between(-thresh, thresh)

        if return_all:
            return z, avg, std, m
        return s.where(m, np.nan)

    def fill_gaps_and_remove_outliers(self, df_total_long, detector_id, which_sample):
        types_vehicle = list(df_total_long['type_vehicle'].drop_duplicates())
        directions = list(df_total_long['direction'].drop_duplicates())
        outliers = []
        self.editor.insert(tk.END, f'Провожу обработку выбросов во временном ряду...\n')

        print('Провожу обработку выбросов во временном ряду...')
        for i in tqdm(types_vehicle):
            for j in directions:
                tmp_df = df_total_long.query(f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)
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
        print('Провожу дополнение временного ряда, заменяю пропуски...')
        for i in tqdm(types_vehicle):
            for j in directions:
                # i = 'Общая интенсивность автомобилей'
                # j = 'Прямое'
                tmp_df = outliers_long.query(f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)

                miss_interval, idx = self.__find_missing_intervals_with_indices(
                    tmp_df.loc[tmp_df.index.isin(self.time_interval_cond)])

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
                    # альтернативный способ через лямбду
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
        if which_sample == 'Rosautodor_1':
            self.coef_sample = pd.read_excel('../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_1')
        elif which_sample == 'Rosautodor_2':
            self.coef_sample = pd.read_excel('../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_2')

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
        max_value.name = 'максимум'
        # Минимум за сутки
        min_value = np.min(imputed_df_day, axis=0)
        min_value.name = 'минимум'
        # Среднее значение за сутки
        mean_value = np.mean(imputed_df_day, axis=0)
        mean_value.name = 'среднесуточное'
        # Коэффициент перехода (максимум делить на среднее) - куда переход? - непонятки
        coeff_trans = max_value / mean_value
        coeff_trans.name = 'коэф_прехода'

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
        avg_annual_per_24_h_TG = avg_annual_per_24_h_TG.groupby('TG')[['Итого', 'Прямое', 'Обратное']].sum() / 365

        # avg_annual_per_24_h_TG = imputed_df_TG.groupby(by=['direction', 'TG'])['Количество_ТГ'].sum() / 365 # а можно было в одну строчку D;    ;-----;

        # global bs_SSID
        # global bs_intensivnosti
        # SSID - транслитерация от СреднеСуточная Интенсивность Движения
        basic_stats_SSID = pd.concat([avg_annual_per_24_h_TG.astype(str).replace('nan', '0.0'),
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
        basic_stats_SSID.index = [detector_id] * len(basic_stats_SSID)

        basic_stats_intensivnosti = pd.concat([coeff_by_month,
                                               coeff_by_weekday,
                                               coeff_by_hour], axis=0).reset_index()
        basic_stats_intensivnosti.index = [detector_id] * len(basic_stats_intensivnosti)

        return basic_stats_SSID, basic_stats_intensivnosti

    def plot_graphs(self, df_main_clear, df_total_long, cur_year, file, freq='d'):
        self.parent.update_idletasks()
        time.sleep(2)
        self.editor.insert(tk.END, f'Рисую графики для файла {file}...\n')
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(20, 10))
        fig.suptitle(f'{file}', fontsize=13)
        colors_line = ['cornflowerblue', 'darkblue']
        colors_dots = ['lime', 'limegreen']
        directions = ['Прямое', 'Обратное']

        for idx, j in enumerate(directions):
            # j = 'Прямое'
            tmp_df = df_main_clear.query(f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'").copy()
            tmp_df.loc[:, 'y_m_d'] = tmp_df.index.strftime('%Y-%m-%d').values
            tmp_df.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df['y_m_d'])
            tmp_df_day = tmp_df.loc[self.time_interval_cond].groupby(by=['y_m_d', 'direction'])[
                'Количество'].sum().unstack(1)

            # tmp_df_raw - этот нужен для отрисовки пропусков в виде маркеров на графике зеленым цветом (где nan, там рисуем метку)
            tmp_df_raw = df_total_long.query(
                f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'").replace(0, np.nan).copy()
            tmp_df_raw.loc[:, 'y_m_d'] = tmp_df_raw.index.strftime('%Y-%m-%d').values
            tmp_df_raw.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df_raw['y_m_d'])

            tmp_df_raw_day = tmp_df_raw.loc[self.time_interval_cond].groupby(by=['y_m_d', 'direction'])[
                'Количество'].sum().unstack(1).replace(0, np.nan)

            if not tmp_df.empty:
                if freq == 'd':
                    # tmp_df_day.plot(style=colors_line[idx % 2], ax=ax1, legend=False, label=f'{j} направление')

                    # tmp_df_day[tmp_df_raw_day.isna().any(axis=1)].reindex(tmp_df_day.index).replace(0, np.nan) \
                    #     .plot(style='o', color=colors_dots[idx % 2], ax=ax1, legend=False, label=f'{j} направление')

                    ax1.plot(tmp_df_day.index, tmp_df_day.replace(np.nan, 0),
                             color=colors_line[idx % 2],
                             label=f'{j} направление',
                             linewidth=0.5, zorder=1)

                    ax1.scatter(tmp_df_day[tmp_df_raw_day.isna().any(axis=1)].reindex(tmp_df_day.index).index,
                                tmp_df_day[tmp_df_raw_day.isna().any(axis=1)].reindex(tmp_df_day.index).replace(0, np.nan),
                                color=colors_dots[idx % 2], marker='o', s=10, label=f'{j} направление', zorder=2)

                    ax1.xaxis.set_major_locator(MonthLocator())
                    ax1.xaxis.set_major_formatter(DateFormatter('%Y-%m'))

                elif freq == 'h':
                    ax1.scatter(tmp_df[tmp_df.index.isin(self.time_interval_cond) & tmp_df_raw.isna().any(axis=1)].loc[:,
                                'Количество'].index,
                                tmp_df[tmp_df.index.isin(self.time_interval_cond) & tmp_df_raw.isna().any(axis=1)].loc[:,
                                'Количество'].replace(0, np.nan),
                                color=colors_dots[idx % 2], marker='o', s=10, label=f'{j} направление', zorder=2)

                    ax1.plot(tmp_df.loc[tmp_df.index.isin(self.time_interval_cond), 'Количество'].index,
                             tmp_df.loc[tmp_df.index.isin(self.time_interval_cond), 'Количество'],
                             color=colors_line[idx % 2],
                             label=f'{j} направление',
                             linewidth=0.5, zorder=1)
                    ax1.xaxis.set_major_locator(MonthLocator())
                    ax1.xaxis.set_major_formatter(DateFormatter('%Y-%m'))

                ax1.set_title(
                    'Суточная интенсивность: скорректированные (дополненные) данные')  # Заголовок первого графика
                ax1.legend(loc='best')  # Указываем, где разместить легенду
                ax1.set_ylabel('Количество / Заполненное значение')  # Подпись оси Y
                ax1.set_xlabel('Время')  # Подпись оси X

        for idx, j in enumerate(directions):
            tmp_df_raw = df_total_long.query(
                f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'").replace(np.nan, 0).copy()
            tmp_df_raw.loc[:, 'y_m_d'] = tmp_df_raw.index.strftime('%Y-%m-%d').values
            tmp_df_raw.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df_raw['y_m_d'])

            tmp_df_raw_day = tmp_df_raw.loc[self.time_interval_cond].groupby(by=['y_m_d', 'direction'])[
                'Количество'].sum().unstack(1)

            if not tmp_df_raw.empty:
                if freq == 'd':
                    # tmp_df_raw_day.plot(style=colors_line[idx % 2], ax=ax2, legend=False, label=f'{j} направление')
                    ax2.plot(tmp_df_raw_day.index,
                             tmp_df_raw_day,
                             color=colors_line[idx % 2],
                             label=f'{j} направление',
                             linewidth=0.5, zorder=1)
                elif freq == 'h':
                    ax2.plot(tmp_df_raw.loc[tmp_df_raw.index.isin(self.time_interval_cond), 'Количество'].index,
                             tmp_df_raw.loc[tmp_df_raw.index.isin(self.time_interval_cond), 'Количество'],
                             color=colors_line[idx % 2],
                             label=f'{j} направление',
                             linewidth=0.5, zorder=1)
                    ax2.xaxis.set_major_locator(MonthLocator())
                    ax2.xaxis.set_major_formatter(DateFormatter('%Y-%m'))

                ax2.set_title('Суточная интенсивность: исходные данные')  # Заголовок второго графика
                ax2.legend(loc='best')  # Указываем, где разместить легенду
                ax2.set_ylabel('Количество')  # Подпись оси Y
                ax2.set_xlabel('Время')  # Подпись оси X

        plt.rcParams['font.size'] = 12
        plt.legend(loc='best')
        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(f'../raw_data/{cur_year}/Графики/' + 'PIC' + file[3:len(file) - 5] + '.png')
        plt.show()
