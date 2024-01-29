"""
Скрипт для создания сопроводительной документации
Основной скрипт
"""
import numpy
import numpy as np

from create_fis_frdo import create_fis_frdo # модуль для создания файла фис фрдо
from decl_case import declension_fio_by_case # функция для склонения фио и создания инициалов
from decl_case import declension_lst_fio_columns_by_case # функция для склонения колонок с фио из листа описания курса
from generate_docs import generate_docs # модуль для создания документов
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

class SameNameColumn(Exception):
    """
    Исключение для обработки случая когда в двух листах есть одинаковые названия колонок
    """
    pass

class SamePathFolder(Exception):
    """
    Исключение для случая когда одна и та же папка выбрана в качестве источника и конечной папки
    """
    pass

class NotFillMainValue(Exception):
    """
    Исключение для проверки заполнения 3 главных параметров: Наименование_программы, Тип_программы, Вид_документа
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
        if folder_template == result_folder:
            raise SamePathFolder


        # Предобработка датафрейма с данными курса
        descr_df = pd.read_excel(data_file, sheet_name='Описание', dtype=str,usecols='A:B')  # получаем данные
        descr_df.dropna(how='all',inplace=True) # удаляем пустые строки
        # траснпонируем
        descr_df = descr_df.transpose()
        descr_df.columns = descr_df.iloc[0] # устанавливаем первую строку в качестве названий колонок
        descr_df.drop(labels='Наименование параметра',inplace=True,axis=0) # удаляем первую строку
        descr_df.index = [0] # переименовываем оставшийся индекс в 0
        # Проверяем наличие колонок
        desc_check_cols = {'Наименование_программы','Тип_программы','Вид_документа','Квалификация_профессия_специальность','Разряд_класс','Разряд_класс_текст','Дата_начало','Дата_конец','Объём',
                           'Руководитель','Руководитель_подразд','Секретарь','Преподаватель','Куратор','База','Председатель_АК'}
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

        lst_check_fill_main_value = list(descr_df.iloc[0,:3])
        if pd.isnull(lst_check_fill_main_value).any():
            raise NotFillMainValue










        # Предобработка датафрейма с данными слушателей
        data_df = pd.read_excel(data_file, sheet_name='Данные физлиц', dtype=str)  # получаем данные
        # Проверяем наличие нужных колонок в файле с данными
        check_columns_data = {'Номер_удостоверения','Рег_номер','Дата_рождения','Пол','СНИЛС','Гражданство','Уровень_образования'
            ,'Серия_паспорта','Номер_паспорта','Кем_выдан_паспорт','Дата_выдачи_паспорта'} # проверяемые колонки
        diff_cols = check_columns_data.difference(set(data_df.columns))
        if len(diff_cols) != 0:
            raise NotNameColumn  # если есть разница вызываем и обрабатываем исключение
        data_df.dropna(how='all',inplace=True) # удаляем пустые строки
        # Обрабатываем вариант создаем доп колонки связанные с ФИО
        data_df = declension_fio_by_case(data_df,result_folder)
        if 'ФИО_представителя' in data_df.columns:
            data_df= declension_lst_fio_columns_by_case(data_df,['ФИО_представителя'])

        # Обрабатываем колонки из датафрейма с описанием курса склоняя по падежам и создавая иницииалы
        descr_fio_cols =['Руководитель','Руководитель_подразд','Секретарь','Преподаватель','Куратор','Председатель_АК'] # список колонок для которых нужно создать падежи и инициалы
        descr_df = declension_lst_fio_columns_by_case(descr_df,descr_fio_cols)




        """
            Конвертируем даты из формата ГГГГ-ММ-ДД в ДД.ММ.ГГГГ
            """
        # делаем строковыми названия колонок
        descr_df.columns = list(map(str,descr_df.columns))
        data_df.columns = list(map(str,data_df.columns))

        # проверяем на совпадение названий колонок в обоих листах
        intersection_columns = set(descr_df.columns).intersection(set(data_df.columns))
        if len(intersection_columns) > 0:
            raise SameNameColumn

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

        # Создаем файл ФИС-ФРДО Если нет папки или файлов то ничего не создаем
        if os.path.exists(f'{folder_template}/ФИС-ФРДО/Шаблон ФИС-ФРДО ДПО.xlsx') and os.path.exists(f'{folder_template}/ФИС-ФРДО/Шаблон ФИС-ФРДО ПО.xlsx'):
            create_fis_frdo(data_df,descr_df,folder_template,result_folder,type_program,descr_df['Вид_документа'].values[0])
        else:
            messagebox.showwarning('Линди Создание документов ДПО,ПО',f'ПРЕДУПРЕЖДЕНИЕ !!!\n В папке {folder_template} не найдена папка ФИС-ФРДО или файлы шаблонов в этой папке.\n'
                                   'В папке ФИС-ФРДО должно быть 2 файла, эти файлы должны иметь название Шаблон ФИС-ФРДО ПО и Шаблон ФИС-ФРДО ДПО.\n'
                                                                'Отсутствие этой папки НЕ ПОВЛИЯЕТ на создание остальных документов.')

        # создаем словари с данными для колонок описания программы

        # получаем списки валидных названий колонок
        descr_valid_cols,descr_not_valid_cols = selection_name_column(list(descr_df.columns),r'^[a-zA-ZЁёа-яА-Я_]+$')
        data_valid_cols, data_not_valid_cols = selection_name_column(list(data_df.columns),r'^[a-zA-ZЁёа-яА-Я_]+$')
        # TODO файл с ошибками и предупреждениями

        # заполняем наны пробелами
        descr_df.fillna(' ',inplace=True)
        data_df.fillna(' ',inplace=True)

        # Словарь с описанием курса
        dct_descr = dict()
        for name_column in descr_valid_cols:
            dct_descr[name_column] = descr_df.loc[0,name_column]
        type_form = 'ФЛ'  # указываем физлицо или юрлицо
        generate_docs(dct_descr,data_df[data_valid_cols],folder_template,result_folder,type_program,type_form)
        messagebox.showinfo('Линди Создание документов ДПО,ПО','Создание документов успешно завершено !')
    except NotNameColumn:
        messagebox.showerror('Линди Создание документов ДПО,ПО',
                             f'В файле {data_file} не найдены следующие колонки {diff_cols}')
    except SameNameColumn:
        messagebox.showerror('Линди Создание документов ДПО,ПО',
                             f'На листе с описанием и на листе со списком найдены одинаковые названия колонок {intersection_columns}\n'
                             f'переименуйте колонки')

    except SamePathFolder:
        messagebox.showerror('Линди Создание документов ДПО,ПО',
                             f'Выбрана одна и та же папка в качесте исходной и конечной.\n'
                             f'Исходная и конечная папки должны быть разными !!!')
    except NotFillMainValue:
        messagebox.showerror('Линди Создание документов ДПО,ПО',
                             f'Заполните значения: Наименование_программы,Тип_программы,\nВид_документа !')
    except PermissionError as e:
        messagebox.showerror('Линди Создание документов ДПО,ПО',
                             f'Закройте файлы созданные программой')



if __name__ == '__main__':
    main_data_file = 'data/Данные по курсу.xlsx'
    # main_data_file = 'data/Данные по курсу несовершеннолетние.xlsx'
    # main_data_file = 'data/Пустая таблица для заполнения курсов.xlsx'
    main_folder_template = 'data/Шаблоны'
    # main_folder_template = 'data/Шаблоны/empty'
    main_result_folder = 'data/Результат'

    create_docs(main_data_file,main_folder_template,main_result_folder)
    print('Lindy Booth !!!')