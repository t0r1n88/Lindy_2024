"""
Скрипт для создания сопроводительной документации
Основной скрипт
"""
from create_fis_frdo import create_fis_frdo # модуль для создания файла фис фрдо
from decl_case import declension_fio_by_case # модуль для склонения фио и создания инициалов
from support_functions import * # вспомогательные функции
import pandas as pd
import openpyxl
from tkinter import messagebox
import os
from datetime import datetime
import re
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None


class NotNameColumn(Exception):
    """
    Исключение для обработки случая когда не совпадают названия колонок
    """
    pass





def create_docs(data_file:str,folder_template:str,result_folder:str):
    """
    Скрипт для сопроводительной документации. Точка входа
    :param data_file: файл Excel с данными
    :param folder_template: папка с шаблонами
    :param result_folder: итоговая папка
    :return: Документация в формате docx и файл ФИс-ФРДО
    """
    try:
        # Предобработка датафрейма с данными курса
        descr_df = pd.read_excel(data_file, sheet_name='Описание', dtype=str,nrows=1)  # получаем данные
        # Проверяем наличие колонок
        desc_check_cols = {'Наименование_программы','Тип_программы','Квалификация_профессия_специальность','Разряд_класс','Дата_начало','Дата_конец','Объем',
                           'ФИО_руководитель','Должность_руководитель','Основание_родит_падеж','ФИО_секретарь','База'}
        diff_cols = desc_check_cols.difference(set(descr_df.columns))
        if len(diff_cols) != 0:
            raise NotNameColumn
        descr_df = descr_df.applymap(lambda x:re.sub(r'\s+',' ',x) if isinstance(x,str) else x) # очищаем от лишних пробелов
        descr_df = descr_df.applymap(lambda x:x.strip() if isinstance(x,str) else x) # очищаем от пробелов в начале и конце

        # Получаем тип программы ДПО или ПО
        dpo_set = {'Повышение квалификации','Профессиональная переподготовка'}
        if descr_df.loc[0,'Тип_программы'] in dpo_set:
            type_program = 'ДПО'
        else:
            type_program = 'ПО'


        # Создаем единичные переменные
        name_program = descr_df.loc[0,'Наименование_программы']
        type_course  = descr_df.loc[0,'Тип_программы']
        name_qval = descr_df.loc[0,'Квалификация_профессия_специальность']
        category = descr_df.loc[0,'Разряд_класс']
        date_begin = descr_df.loc[0,'Дата_начало']
        date_end = descr_df.loc[0,'Дата_конец']
        volume = descr_df.loc[0,'Объем']
        fio_chief = descr_df.loc[0,'ФИО_руководитель']
        chief_position = descr_df.loc[0,'Должность_руководитель']
        name_doc_rod_case = descr_df.loc[0,'Основание_родит_падеж']
        fio_secretary = descr_df.loc[0,'ФИО_секретарь']
        base = descr_df.loc[0,'База']

        # Предобработка датафрейма с данными слушателей
        data_df = pd.read_excel(data_file, sheet_name='Данные', dtype=str)  # получаем данные
        # Проверяем наличие нужных колонок в файле с данными
        check_columns_data = {'Номер_удостоверения','Рег_номер','Дата_рождения','Пол','СНИЛС','Гражданство','Уровень_образования'
            ,'Серия_паспорта','Номер_паспорта','Кем_выдан_паспорт','Дата_выдачи_паспорта'} # проверяемые колонки
        diff_cols = check_columns_data.difference(set(data_df.columns))
        if len(diff_cols) != 0:
            raise NotNameColumn  # если есть разница вызываем и обрабатываем исключение
        # Обрабатываем вариант создаем доп колонки связанные с ФИО
        data_df = declension_fio_by_case(data_df)
        """
            Конвертируем даты из формата ГГГГ-ММ-ДД в ДД.ММ.ГГГГ
            """
        # делаем строковыми названия колонок
        descr_df.columns = list(map(str,descr_df.columns))
        data_df.columns = list(map(str,data_df.columns))
        # Обрабатываем колонки с датами в описании
        lst_date_columns_descr = []  # список для колонок с датами
        for idx, column in enumerate(descr_df.columns):
            if 'дата' in column.lower():
                lst_date_columns_descr.append(idx)

        descr_df = convert_string_date(descr_df,lst_date_columns_descr)

        # обрабатываем колонки с датами в списке
        lst_date_columns_data = []  # список для колонок с датами
        for idx, column in enumerate(data_df.columns):
            if 'дата' in column.lower():
                lst_date_columns_data.append(idx)
        data_df = convert_string_date(data_df,lst_date_columns_data)

        # data_df['Дата_рождения'] = data_df['Дата_рождения'].apply(convert_date_yandex)
        # data_df['Дата_выдачи_паспорта'] = data_df['Дата_выдачи_паспорта'].apply(convert_date_yandex)

        # Создаем файл ФИС-ФРДО
        create_fis_frdo(data_df,descr_df,folder_template,result_folder,type_program,data_file)

        # создаем словари с данными для колонок описания программы

        # получаем списки валидных названий колонок
        descr_valid_cols,descr_not_valid_cols = selection_name_column(list(descr_df.columns),r'^[a-zA-ZЁёа-яА-Я_]+$')
        print(descr_valid_cols)
        print(descr_not_valid_cols)
        data_valid_cols, data_not_valid_cols = selection_name_column(list(data_df.columns),r'^[a-zA-ZЁёа-яА-Я_]+$')
        print(data_valid_cols)
        print(data_not_valid_cols)



        # Создаем словари
        # Словарь с описанием курса
        dct_descr = dict()
        for name_column in descr_valid_cols:
            dct_descr[name_column] = descr_df.loc[0,name_column]
        print(dct_descr)













        # data_df.to_excel('data/Результат/dasd.xlsx',index=False,header=True)
        # descr_df.to_excel('data/Результат/Исходник Описание.xlsx',index=False,header=True)



    except NotNameColumn:
        messagebox.showerror('Создание документов ДПО,ПО',
                             f'В файле {data_file} не найдены следующие колонки {diff_cols}')


if __name__ == '__main__':
    main_data_file = 'data/Таблица для заполнения бланков.xlsx'
    main_folder_template = 'data/Шаблоны'
    main_result_folder = 'data/Результат'

    create_docs(main_data_file,main_folder_template,main_result_folder)
    print('Lindy Booth !!!')