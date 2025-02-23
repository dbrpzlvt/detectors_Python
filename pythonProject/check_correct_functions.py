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


cur_year = 2024
file = 'PRE_км 198+500 а.д М-10 Россия Москва – Санкт-Петербург .xlsx'
directory_pre = f'../raw_data/{str(cur_year)}/Первичная обработка/'
wb = openpyxl.load_workbook(os.path.join(directory_pre, file))

sheet = wb.active
columnA = sheet['A']
last_row = len(sheet['A'])

wb_data = xw.Book(os.path.join(directory_pre, file))

road_name = wb_data.sheets[0]['A6'].value
road_name = road_name.replace('/', '.')  # Слеш заменяю на точку
road_name = road_name.replace('"', '')  # Удаляю символы кавычек
road_name = road_name.replace('?', '')  # вопросительный знак
road_name = road_name.replace('!', '')  # восклицательный знак
road_name = road_name.replace('*', '')  # символ звёздочки
road_name = road_name.replace('\n', '')  # Убираю символы переноса строки
road_name = road_name.replace('\t', '')  # Убираю символы табуляции, которые могут мешать сохранению файла

my_road_values = wb_data.sheets[0].range('A5:A6').options(ndim=2).value

# wb_sample = xw.Book('../sample_fda_var1.xlsm')
# wb_sample.sheets['Исходные данные'].range('A3:A4').value = my_road_values

my_values = wb_data.sheets[0].range('A7:AB' + str(last_row)).options(ndim=2).value
# wb_sample.sheets['Исходные данные'].range('A5:AB' + str(last_row - 2)).value = my_values

df = pd.DataFrame(my_values)

if list(df.tail(3)[0])[0] == 'Итого':
    df.drop(df.tail(3).index, inplace=True)  # удаляю последние три строки ['Итого', 'Среднее', '%']
elif list(df.tail(3)[0])[0] != 'Итого' and list(df.tail(3)[0])[1] == 'Среднее':
    df.drop(df.tail(2).index, inplace=True)
elif list(df.tail(3)[0])[0] != 'Итого' and list(df.tail(3)[0])[1] != 'Среднее' and list(df.tail(3)[0])[2] == '%':
    df.drop(df.tail(1).index, inplace=True)
else:
    pass

column_names = wb_data.sheets['Исходные данные'].range('A3:AB4').value
column_names = pd.DataFrame(column_names).ffill(axis=1).ffill(axis=0)
df.columns = pd.MultiIndex.from_arrays(column_names[:2].values, names=['type_vehicle', 'direction'])
df[('Дата', 'Дата')] = pd.to_datetime(df[('Дата', 'Дата')], format='%d.%m.%Y %H:%M:%S')
df.iloc[:, 1:] = df.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

# принудительно делаем из object - float64
for col in df.columns[1:]:
    if isinstance(df[col].dtype, object):
        df[col] = df[col].astype(float)


def add_previous_december_data():
    previous_year_files = os.listdir(f'../raw_data/{str(cur_year - 1)}/Первичная обработка')
    wb_prev_yr_file = pd.DataFrame()
    for i in previous_year_files:
        # print(i)
        if file == i:
            print('Файл за предыдущий год найден! ' + i)
            wb_prev_yr_file = pd.read_excel(
                os.path.join(f'../raw_data/{str(cur_year - 1)}/Первичная обработка/', file)).iloc[:, 0:-1]
            wb_prev_yr_file.columns = pd.MultiIndex.from_arrays(column_names[:2].values,
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
            # return wb_prev_yr_file
    return wb_prev_yr_file


tmp_df = add_previous_december_data()

list_dfs = [df, tmp_df]
if not tmp_df.empty:
    df = pd.concat([i.dropna(axis=1, how='all') for i in list_dfs]).reset_index(drop=True)


# df.iloc[:, 1:] = df.iloc[:, 1:].astype(float)
# df.dtypes
# df = df.rename(columns={('Дата', 'Дата'): "Дата"})

# Достраиваю временной ряд, если в исходных данных полностью отсутствуют дни/часы замеров
full_time_series = pd.DataFrame({('Дата', 'Дата'):
                                     pd.date_range('2023-12-01 00:59:59', '2025-01-31 23:59:59', freq='1h')})
full_time_series = pd.to_datetime(full_time_series[('Дата', 'Дата')], format='%d.%m.%Y %H:%M:%S')
df = full_time_series.to_frame().merge(df, how='left', on=[('Дата', 'Дата')])
df.fillna(0, inplace=True)
df.set_index(('Дата', 'Дата'), inplace=True)
# df.dtypes
time_interval_cond = full_time_series[('2024-01-01 00:59:59' < full_time_series) & (full_time_series < '2024-12-31 23:59:59')]

max_measured = 8784
fact_measured = sum(df[('Общая интенсивность автомобилей', 'Итого')][time_interval_cond].replace(0, np.nan).notna())
fullness = fact_measured/max_measured


def check_correct_data():
    # df.columns[[6, 9, 12, 15, 18, 21]]
    sum_non_passenger_cars = pd.DataFrame({('Дата', 'Дата'): {},
                                           ('Сумма не-легковых', 'Итого'): {},
                                           ('Сумма не-легковых', 'Прямое'): {},
                                           ('Сумма не-легковых', 'Обратное'): {}})
    sum_non_passenger_cars['Дата', 'Дата'] = df.index
    sum_non_passenger_cars.set_index(('Дата', 'Дата'), inplace=True)
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

    error_amount_direction = pd.DataFrame({('Дата', 'Дата'): {},
                                           ('Величина ошибки', 'Прямое'): {},
                                           ('Величина ошибки', 'Обратное'): {}})
    error_amount_direction['Дата', 'Дата'] = df.index
    error_amount_direction.set_index(('Дата', 'Дата'), inplace=True)
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

    checker_data = pd.DataFrame({('Дата', 'Дата'): {},
                                           ('Величина ошибки', 'Прямое'): {},
                                           ('Величина ошибки', 'Обратное'): {}})
    checker_data['Дата', 'Дата'] = df.index
    checker_data.set_index(('Дата', 'Дата'), inplace=True)
    checker_data['Величина ошибки', 'Прямое'] = np.where(
        (df['Общая интенсивность автомобилей', 'Прямое'] < 0.25 * df['Общая интенсивность автомобилей', 'Обратное']) |
        (df.sum(axis=1) == 0) |  # .loc[(df.index > '2024-01-11 11:59:59') & (df.index < '15.01.2024 16:59:59')]
        (sum_non_passenger_cars['Сумма не-легковых', 'Прямое'] > 10 * (df['Общая интенсивность автомобилей', 'Прямое'] + df['Общая интенсивность автомобилей', 'Обратное'])), 'Данные НЕкорректны', 'Данные корректны')

    checker_data['Величина ошибки', 'Обратное'] = np.where(
        (df['Общая интенсивность автомобилей', 'Обратное'] < 0.25 * df['Общая интенсивность автомобилей', 'Прямое']) |
        (df.sum(axis=1) == 0) |
        (sum_non_passenger_cars['Сумма не-легковых', 'Обратное'] > 10 * (df['Общая интенсивность автомобилей', 'Прямое'] + df['Общая интенсивность автомобилей', 'Обратное'])), 'Данные НЕкорректны', 'Данные корректны')

    print('Некорректные данные: прямое направление: '+str(sum(np.where(checker_data['Величина ошибки', 'Прямое'] == 'Данные НЕкорректны', 1, 0))))
    print('Некорректные данные: обратное направление: '+str(sum(np.where(checker_data['Величина ошибки', 'Обратное'] == 'Данные НЕкорректны', 1, 0))))
    # checker_data.reset_index(inplace=True)

    return checker_data


checker_data = check_correct_data()

df_long = df.melt(ignore_index=False, value_name='Количество').reset_index().rename(columns={('Дата', 'Дата'): 'Дата', 'variable_0': 'type_vehicle', 'variable_1': 'direction'})
checker_data_long = checker_data.melt(ignore_index=False, col_level=1, value_vars=['Обратное', 'Прямое'], var_name='direction', value_name='Корректность').reset_index().rename(columns={('Дата', 'Дата'): "Дата"})

df_total_long = df_long.merge(checker_data_long, how='left', on=['Дата', 'direction'])  # .rename(columns={'Значение_x': 'Количество', 'Значение_y': 'Корректность'})
df_total_long = df_total_long[(df_total_long.type_vehicle != 'Загрузка, %') & (df_total_long.type_vehicle != 'Скорость, км/ч')]

df_total_long.set_index('Дата', inplace=True)


def find_missing_intervals_with_indices(df_in):
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


def fill_missing_values(df_in, date_time_idx, max_depth=10):
    """
    Функция для заполнения пропущенных значений с помощью простого среднего.
    Относительно NaN ищется значение неделю вперед/назад и от них берется среднее.
    Недостаток: ряд может иметь переменную сезонность (новогодние праздники, закрытие дороги, ремонт и т.п.).
    Из-за этого среднее может быть не совсем корректно.
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

    # Если не удалось найти значения, возвращаем NaN
    return np.nan


# удаляю выбросы с помощью z-оценки
def zscore(s, window, thresh=3, return_all=False):
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

# ящики с усами чтобы понять где выбросы
# import seaborn as sns
#
# sns.boxplot(data=df_total_long.query(f"type_vehicle == 'Общая интенсивность автомобилей' and direction == 'Прямое'").replace(0, np.nan),
#             x=df_total_long.query(f"type_vehicle == 'Общая интенсивность автомобилей' and direction == 'Прямое'").replace(0, np.nan).index,
#             y='Количество')


types_vehicle = list(df_total_long['type_vehicle'].drop_duplicates())
directions = list(df_total_long['direction'].drop_duplicates())
outliers = []
print('Провожу обработку выбросов во временном ряду...')
for i in tqdm(types_vehicle):
    for j in directions:
        tmp_df = df_total_long.query(f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)
        if np.isnan(tmp_df.loc[:, 'Количество']).all():
            tmp_df.loc[:, "Количество"] = 0
            outliers.append(tmp_df)
            continue
        tmp_df.loc[:, 'Количество'] = zscore(tmp_df['Количество'], '24h')
        outliers.append(tmp_df)

outliers_long = pd.concat(outliers)


imputed = []
print('Провожу дополнение временного ряда, заменяю пропуски...')
for i in tqdm(types_vehicle):
    for j in directions:
        # i = 'Общая интенсивность автомобилей'
        # j = 'Прямое'
        tmp_df = outliers_long.query(f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)
        # tmp_df = df_total_long.query(
        #     f"type_vehicle == '{i}' and direction == '{j}'").replace(0, np.nan)
        miss_interval, idx = find_missing_intervals_with_indices(
            tmp_df.loc[tmp_df.index.isin(time_interval_cond)])

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
            # last_index = tmp_df.index.get_loc(end)  # Получаем позицию конца интервала
            # first_index = tmp_df.index.get_loc(start)  # Получаем позицию начала интервала

            # if first_index - 1 < len(tmp_df):  # Проверяем, что предыдущий индекс существует
            #     tmp_df.iloc[first_index - 1, tmp_df.columns.get_loc('Количество')] = np.nan
            #
            # if last_index + 1 < len(tmp_df):  # Проверяем, что следующий индекс существует
            #     tmp_df.iloc[last_index + 1, tmp_df.columns.get_loc('Количество')] = np.nan
        if np.isnan(tmp_df.loc[:, 'Количество']).all():
            tmp_df.loc[:, "Количество"] = 0
            imputed.append(tmp_df)
            continue
            # альтернативный способ через лямбду
        tmp_df.loc[:, "Количество"] = tmp_df.loc[:, 'Количество'].index.map(
            lambda dt: fill_missing_values(tmp_df.loc[:, 'Количество'], dt) if np.isnan(
                tmp_df.at[dt, 'Количество']) else tmp_df.at[dt, 'Количество'])  # else df_total_long.loc[cond].at[dt, 'Количество'])
        imputed.append(tmp_df)

imputed_df = pd.concat(imputed)  # датафрейм с дополненными данными

which_sample = 'Rosautodor_1'
if which_sample == 'Rosautodor_1':
    coef_sample = pd.read_excel('../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_1')
elif which_sample == 'Rosautodor_2':
    coef_sample = pd.read_excel('../raw_data/coeff_transform_to_TG.xlsx', sheet_name='sample_2')

imputed_df = imputed_df.loc[time_interval_cond].reset_index().merge(coef_sample, how='left', on='type_vehicle').set_index('Дата')
# imputed_df['TG'] = imputed_df['TG'].fillna('Все ТГ')
imputed_df['Количество_ТГ'] = np.prod([imputed_df['Количество'], imputed_df['coeff']], axis=0)

# tmp = imputed_df.groupby(by=['direction', 'TG', 'Дата'])['Количество_ТГ'].sum().unstack([0, 1])
imputed_df['month'] = imputed_df.index.strftime("%B")
imputed_df['day_of_week'] = imputed_df.index.strftime("%A")
imputed_df['hour'] = imputed_df.index.strftime("%H")
imputed_df['y_m_d'] = imputed_df.index.strftime('%Y-%m-%d')
imputed_df['y_m_d'] = pd.to_datetime(imputed_df['y_m_d'])
# imputed_df.dtypes
foo = imputed_df.groupby(by=['y_m_d', 'direction', 'TG'])['Количество_ТГ'].sum().unstack([1, 2])

max_value = np.max(foo, axis=0).replace(0, np.nan)
min_value = np.min(foo, axis=0)
mean_value = np.mean(foo, axis=0)
coeff_trans = max_value / mean_value

# создаю новый series для хранения результата (день, где был максимум)
days_of_max = pd.Series(index=foo.columns, dtype=object)

for idx in foo.columns:
    # idx = foo.columns[0]
    col = idx
    matches = foo[col].isin([max_value[idx]])  # проверяю совпадения с соответствующим значением из max_value

    if matches.sum() == 1:  # сохраняю день
        days_of_max[idx] = foo[idx].index[matches].values
    elif matches.sum() >= 2:  # если совпадений два или более, сохраняю список дней
        days_of_max[idx] = [foo[idx].index[matches].values]
    else:
        days_of_max[idx] = np.nan  # проставляю np.nan, если совпадений нет

# считаю коэффициенты неравномерности
coeff_by_month = imputed_df.groupby(by=['month', 'direction', 'TG'])['Количество_ТГ'].sum().unstack([1, 2])
coeff_by_month = coeff_by_month.div(
    coeff_by_month.sum(axis=0),
    axis=1)

coeff_by_weekday = imputed_df.groupby(by=['day_of_week', 'direction', 'TG'])['Количество_ТГ'].sum().unstack([1, 2])
coeff_by_weekday = coeff_by_weekday.div(
    coeff_by_weekday.sum(axis=0),
    axis=1)

coeff_by_hour = imputed_df.groupby(by=['hour', 'direction', 'TG'])['Количество_ТГ'].sum().unstack([1, 2])
coeff_by_hour = coeff_by_hour.div(
    coeff_by_hour.sum(axis=0),
    axis=1)

amount_per_year_vehicle = imputed_df.groupby(by=['direction', 'type_vehicle'])['Количество_ТГ'].sum()
avg_annual_per_24_h_TG = imputed_df.groupby(by=['direction', 'TG'])['Количество_ТГ'].sum()/365



# %% способы обработки пропусков
        # альтернативный способ через цикл
        # miss_interval, idx = find_missing_intervals_with_indices(
        #     tmp_df.loc[tmp_df.index.isin(time_interval_cond)])
        # for k, l in zip(miss_interval, idx):
        #     # start, end = l
        #     # print(i, j)
        #     # print(start, end)
        #     for date_time_idx in pd.date_range(l[0], l[1], freq='h'):
        #         # date_time_idx = pd.Timestamp('2024-01-15 15:59:59')
        #         # dt = date_time_idx
        #         found_value = fill_missing_values(tmp_df.loc[:, 'Количество'], date_time_idx)
        #         tmp_df.loc[tmp_df.index == date_time_idx, 'filled_value_new'] = found_value
        #     new_cheburek.append(tmp_df)


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
#         ax = df_imputed.plot(linestyle="-", marker=".", figsize=[10, 5], legend=None)
#         ax = df_imputed[tmp_df.loc[:, 'Количество'].isnull()].plot(ax=ax, legend=None, marker=".", color="r")

# cond = (df_total_long['direction'] == 'Прямое') & (df_total_long['type_vehicle'] == 'Общая интенсивность автомобилей')
# df_total_long.loc[cond, 'Количество'] = df_total_long.loc[cond, 'Количество'].replace(0, np.nan)

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
# %% конец способов обработки пропусков

# создаю датафрейм для ДОПОЛНЕННЫХ данных только с данными по "Общей интенсивности автомобилей"
df_main_clear = []
for j in directions:
    for i in imputed:
        tmp_df = i.query(f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'")
        if not tmp_df[(tmp_df.type_vehicle == 'Общая интенсивность автомобилей')
                      & (tmp_df.direction == j)].empty:
            df_main_clear.append(i)
df_main_clear = pd.concat(df_main_clear)
# df_total_long = df_total_long.query(f"type_vehicle == 'Общая интенсивность автомобилей'")

# for window in range(3):


fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(20, 10))
fig.suptitle(f'{file}', fontsize=13)
colors_line = ['cornflowerblue', 'darkblue']
colors_dots = ['lime', 'limegreen']
directions = ['Прямое', 'Обратное']

for idx, j in enumerate(directions):
    # j = 'Прямое'
    tmp_df = df_main_clear.query(f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'")
    tmp_df.loc[:, 'y_m_d'] = tmp_df.index.strftime('%Y-%m-%d').values
    tmp_df.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df['y_m_d'])

    tmp_df = tmp_df.loc[time_interval_cond].groupby(by=['y_m_d', 'direction'])[
        'Количество'].sum().unstack(1)

    # tmp_df_raw - этот нужен для отрисовки пропусков на графике красным цветом (где nan, там рисуем метку)
    tmp_df_raw = df_total_long.query(
        f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'").replace(0, np.nan)
    tmp_df_raw.loc[:, 'y_m_d'] = tmp_df_raw.index.strftime('%Y-%m-%d').values
    tmp_df_raw.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df_raw['y_m_d'])

    tmp_df_raw = tmp_df_raw.loc[time_interval_cond].groupby(by=['y_m_d', 'direction'])[
        'Количество'].sum().unstack(1).replace(0, np.nan)

    if not tmp_df.empty:
        tmp_df.plot(style=colors_line[idx % 2], ax=ax1, legend=False, label=f'{j} направление')

        tmp_df[tmp_df_raw.isna().any(axis=1)].replace(0, np.nan) \
            .plot(style='o', color=colors_dots[idx % 2], ax=ax1, legend=False, label=f'{j} направление')

        # tmp_df.loc[tmp_df.index.isin(time_interval_cond), 'Количество']\
        #     .plot(style=colors_line[idx % 2], ax=ax1, legend=False, label=f'{j} направление')
        #
        # tmp_df[tmp_df.index.isin(time_interval_cond) & tmp_df_raw.isna().any(axis=1)].loc[:, 'Количество'] \
        #     .plot(style='o', color=colors_dots[idx % 2], ax=ax1, legend=False, label=f'{j} направление')
        ax1.set_title('Суточная интенсивность: скорректированные (дополненные) данные')  # Заголовок первого графика
        ax1.legend(loc='best')  # Указываем, где разместить легенду
        # ax1.xaxis.set_major_locator(MonthLocator())
        # ax1.xaxis.set_major_formatter(DateFormatter('%Y-%m'))
        ax1.set_ylabel('Количество / Заполненное значение')  # Подпись оси Y
        ax1.set_xlabel('Время')  # Подпись оси X

for idx, j in enumerate(directions):
    tmp_df_raw = df_total_long.query(f"type_vehicle == 'Общая интенсивность автомобилей' and direction == '{j}'").replace(np.nan, 0)
    tmp_df_raw.loc[:, 'y_m_d'] = tmp_df_raw.index.strftime('%Y-%m-%d').values
    tmp_df_raw.loc[:, 'y_m_d'] = pd.to_datetime(tmp_df_raw['y_m_d'])

    tmp_df_raw = tmp_df_raw.loc[time_interval_cond].groupby(by=['y_m_d', 'direction'])[
        'Количество'].sum().unstack(1)

    if not tmp_df_raw.empty:
        # ax2.plot(tmp_df.loc[tmp_df.index.isin(time_interval_cond)].index,
        #          tmp_df.loc[tmp_df.index.isin(time_interval_cond), ['Количество']])
        tmp_df_raw.plot(style=colors_line[idx % 2], ax=ax2, legend=False, label=f'{j} направление')
        ax2.set_title('Суточная интенсивность: исходные данные')  # Заголовок второго графика
        ax2.legend(loc='best')  # Указываем, где разместить легенду
        # ax2.xaxis.set_major_locator(MonthLocator())
        # ax2.xaxis.set_major_formatter(DateFormatter('%Y-%m'))
        ax2.set_ylabel('Количество')  # Подпись оси Y
        ax2.set_xlabel('Время')  # Подпись оси X

plt.rcParams['font.size'] = 14
plt.legend(loc='best')
plt.tight_layout(rect=[0, 0.03, 1, 0.95])
# plt.savefig('../raw_data/2024/Графики/' + 'PIC' + file[3:len(file) - 5] + '.png')
plt.show()







df_total_long.loc[cond &
                  (df_total_long.index.isin(time_interval_cond))].plot(style=['k--', 'bo-', 'r*'], figsize=(20, 10))