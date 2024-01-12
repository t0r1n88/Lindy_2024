"""
Модуль для создания файла ФИС ФРДО
"""
import pandas as pd
import openpyxl
from tkinter import messagebox
import os
from datetime import datetime

class NotFileTemplateDPO(Exception):
    """
    Класс для ошибки когда не найден файл шаблона ДПО
    """
    pass

class NotFileTemplatePO(Exception):
    """
    Класс для ошибки когда не найден файл шаблона ПО
    """
    pass

class NotNameColumn(Exception):
    """
    Исключение для обработки случая когда не совпадают названия колонок
    """
    pass

def convert_date_yandex(value:str):
    """
    Функция для конвертации дат из яндекс файла
    :param value: строка
    :return:дата
    """
    date_object = datetime.strptime(value, "%Y-%m-%d") # делаем объект datetime
    return date_object.strftime("%d.%m.%Y") # преобразуем в нужный формат




def write_data_fis_frdo(template_fis_frdo_dpo:openpyxl.Workbook,dct_df:dict,dct_number_column:dict)->openpyxl.Workbook:
    """
    Функция для записи данных в шаблон ФИС -ФРДО
    :param template_fis_frdo_dpo: шаблон ФИС-ФРДО
    :param dct_df: словарь с данными вида -название колонки:список данных в колонке
    :param dct_number_column: словарь вида -название колонки: порядковый номер колонки куда надо записывать данные
    :return: заполненый шаблон
    """
    for name_column,number_col in dct_number_column.items():
        # перебираем словарь с порядковыми номерами колонок
        start_row = 2  # строка с которой будет начинаться записи
        for value in dct_df[name_column]: # записываем данные из словаря с данными
            template_fis_frdo_dpo['Шаблон'].cell(row=start_row, column=number_col, value=value)
            start_row += 1

    return template_fis_frdo_dpo



def create_fis_frdo(data_file:str,folder_template:str,result_folder:str,type_program:str):
    """
    Функция для создания файлов ФИС ФРДО
    :param data_file: файл с данными
    :param result_folder: путь к конечной папке
    :param folder_template:путь к папке с шаблонами
    :param type_program: тип создаваемого файла - ПК или ПО
    :return:файл Excel
    """
    try:
        if type_program == 'ДПО':
            df = pd.read_excel(data_file,sheet_name='Данные', dtype=str) # получаем данные
            df['Дата_рождения'] = df['Дата_рождения'].apply(convert_date_yandex)

            dct_df = df.to_dict(orient='list') # превращаем в словарь где ключ это название колонки а значение это список
            # Создаем словарь для хранения номеров колонок для каждого названия
            dct_number_column = {'Номер_удостоверения':7,'Рег_номер':9,'Фамилия':22,
                                 'Имя':23,'Отчество':24,
                                 'Дата_рождения':25,'Пол':26,'СНИЛС':27}
            # проверяем наличие соответствующих колонок
            diff_cols = set(dct_number_column.keys()).difference(set(dct_df.keys()))
            if len(diff_cols) != 0:
                raise NotNameColumn # если есть разница вызываем и обрабатываем исключение

            template_fis_frdo_dpo = openpyxl.load_workbook(f'{folder_template}/ФИС-ФРДО/Шаблон ФИС-ФРДО ДПО.xlsx')
            fis_frdo_dpo = write_data_fis_frdo(template_fis_frdo_dpo,dct_df,dct_number_column) # Записываем в шаблон
            fis_frdo_dpo.save(f'{result_folder}/ФИС-ФРДО ДПО.xlsx')


    except FileNotFoundError:
        messagebox.showerror('Создание документов ДПО,ПО',
                             f'В папке {folder_template}/ФИС-ФРДО не найден файл шаблона ФИС-ФРДО.\n'
                                                              f'Файлы должны иметь название - Шаблон ФИС-ФРДО ДПО и Шаблон ФИС-ФРДО ПО')

    except NotNameColumn:
        messagebox.showerror('Создание документов ДПО,ПО',
                             f'В файле {data_file} не найдены следующие колонки {diff_cols}')

















if __name__ == '__main__':
    main_data_file = 'data/Таблица для заполнения бланков.xlsx'
    main_folder_template = 'data/Шаблоны'
    main_result_folder = 'data/Результат'
    main_type_program = 'ДПО'

    create_fis_frdo(main_data_file,main_folder_template,main_result_folder,main_type_program)
    print('Lindy Booth !!!')

