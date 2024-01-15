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

class NotNameColumn(Exception):
    """
    Исключение для обработки случая когда не совпадают названия колонок
    """
    pass


def create_docs(data_file:str,folder_template:str,result_folder:str,type_program:str):
    """
    Скрипт для сопроводительной документации. Точка входа
    :param data_file: файл Excel с данными
    :param folder_template: папка с шаблонами
    :param result_folder: итоговая папка
    :param type_program: тип программы ДПО или ПО
    :return: Документация в формате docx и файл ФИс-ФРДО
    """
    try:
        # Предобработка датафрейма с данными курса
        descr_df = pd.read_excel(data_file, sheet_name='Описание', dtype=str)  # получаем данные

        # Предобработка датафрейма с данными слушателей
        data_df = pd.read_excel(data_file, sheet_name='Данные', dtype=str)  # получаем данные
        # Проверяем наличие нужных колонок в файле с данными
        check_columns_data = {'Номер_удостоверения','Рег_номер','Дата_рождения','Пол','СНИЛС','Гражданство','Уровень_образования'
            ,'Серия_паспорта','Номер_паспорта','Кем_выдан_паспорт','Дата_выдачи_паспорта'} # проверяемые колонки
        diff_cols = check_columns_data.difference(set(data_df.columns))
        if len(diff_cols) != 0:
            raise NotNameColumn  # если есть разница вызываем и обрабатываем исключение
        # Обрабатываем вариант создаем доп колонки связанные с ФИО


        """
            Конвертируем даты из формата ГГГГ-ММ-ДД в ДД.ММ.ГГГГ
            """
        data_df['Дата_рождения'] = data_df['Дата_рождения'].apply(convert_date_yandex)
        data_df['Дата_выдачи_паспорта'] = data_df['Дата_выдачи_паспорта'].apply(convert_date_yandex)



    except NotNameColumn:
        messagebox.showerror('Создание документов ДПО,ПО',
                             f'В файле {data_file} не найдены следующие колонки {diff_cols}')


if __name__ == '__main__':
    main_data_file = 'data/Таблица для заполнения бланков.xlsx'
    main_folder_template = 'data/Шаблоны'
    main_result_folder = 'data/Результат'
    main_type_program = 'ДПО'

    create_docs(main_data_file,main_folder_template,main_result_folder,main_type_program)
    print('Lindy Booth !!!')