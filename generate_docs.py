"""
Скрипт для генерации документов
"""
import pandas as pd
import numpy as np
import os
import shutil
from dateutil.parser import ParserError
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document
from docx2pdf import convert
from tkinter import messagebox
from jinja2 import exceptions
import time
import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
import platform
import logging
import tempfile
import re

logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)

class NotFileSource(Exception):
    """
    Исключение для обработки случая когда не найдены файлы внутри исходной папки
    """
    pass


def combine_all_docx(filename_master, files_lst,path_to_end_folder_doc):
    """
    Функция для объединения файлов Word взято отсюда
    https://stackoverflow.com/questions/24872527/combine-word-document-using-python-docx
    :param filename_master: базовый файл
    :param files_list: список с созданными файлами
    :return: итоговый файл
    """

    # Получаем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    number_of_sections = len(files_lst)
    # Открываем и обрабатываем базовый файл
    master = Document(filename_master)
    composer = Composer(master)
    # Перебираем и добавляем файлы к базовому
    for i in range(0, number_of_sections):
        doc_temp = Document(files_lst[i])
        composer.append(doc_temp)
    # Сохраняем файл
    composer.save(f"{path_to_end_folder_doc}/ОБЩИЙ файл от {current_time}.docx")


def copy_folder_structure(source_folder:str,destination_folder:str):
    """
    Функция для копирования структуры папок внутри выбраной папки
    :param source_folder: Исходная папка
    :param destination_folder: конечная папка
    :return: Структура папок как в исходной папке
    """
    # Получаем список папок внутри source_folder

    lst_subdirs =  [] # список для подпапок
    lst_files = [] # список для файлов
    lst_source_folders = [] # список для хранения путей к папкам в исходной папке

    for dirname, dirnames, filenames in os.walk(source_folder):
        # print path to all subdirectories first.
        for subdirname in dirnames:
            lst_subdirs.append(subdirname)
            lst_source_folders.append(f'{dirname}/{subdirname}')

    # ищем файлы
    for dirname, dirnames, filenames in os.walk(source_folder):
        for file in filenames:
            lst_files.append(file)

    # заменяем папку назначения
    lst_dest_folders = [path.replace(source_folder,destination_folder) for path in lst_source_folders]
    for path_folder in lst_dest_folders:
        if not os.path.exists(path_folder):
            os.makedirs(path_folder)
    # создаем словарь где ключ это путь к папкам в исходном файле а значение это путь к папкам в конечной папке
    # проверяем количество найденных папок
    if len(lst_subdirs) != 0:
        dct_path = dict(zip(lst_source_folders,lst_dest_folders))
    else:
        # если подпапок нет то сохраняем в итоговую папку
        dct_path = {source_folder:destination_folder}

    if len(lst_files) == 0:
        raise NotFileSource

    return dct_path


def generate_docs(dct_descr:dict,data_df:pd.DataFrame,source_folder:str,destination_folder:str,type_program:str):
    """
    Основная функция генерации документов
    :param dct_descr: словарь с описанием курса
    :param data_df: датафрейм с данными слушателей
    :param source_folder: исходная папка
    :param destination_folder: конечная папка
    :param type_program: тип программы ДПО или ПО
    :return: Сопроводительная документация в формате docx
    """
    used_name_file = set()  # множество для уже использованных имен файлов
    # Словарь для получения длинных выражений по типу программы
    where_type_program = {'Повышение квалификации':'дополнительную профессиональную программу повышения квалификации',
                          'Профессиональная переподготовка':'дополнительную профессиональную программу профессиональной переподготовки',
                          'Программа профессиональной подготовки по профессии рабочего, должности служащего':'основную программу профессионального обучения профессиональной подготовки',
                          'Программа переподготовки рабочих, служащих':'основную программу профессионального обучения профессиональной переподготовки',
                          'Программа повышения квалификации рабочих, служащих':'основную программу профессионального обучения повышения квалификации рабочих, служащих',}
    dct_descr['Программа_куда'] = where_type_program[dct_descr['Тип_программы']]

    # словарь для родительного падежа
    rod_type_program = {'Повышение квалификации': 'дополнительной профессиональной программы повышения квалификации',
                          'Профессиональная переподготовка': 'дополнительной профессиональной программы профессиональной переподготовки',
                          'Программа профессиональной подготовки по профессии рабочего, должности служащего': 'основной программы профессионального обучения профессиональной подготовки',
                          'Программа переподготовки рабочих, служащих': 'основной программы профессионального обучения профессиональной переподготовки',
                          'Программа повышения квалификации рабочих, служащих': 'основной программы профессионального обучения повышения квалификации рабочих, служащих', }
    dct_descr['Программа_чего'] = rod_type_program[dct_descr['Тип_программы']]

    # словарь для сокращения типов программ
    abbr_type_program = {'Повышение квалификации': 'ДПП ПК',
                          'Профессиональная переподготовка': 'ДПП ПП',
                          'Программа профессиональной подготовки по профессии рабочего, должности служащего': 'ОП ПО ПП',
                          'Программа переподготовки рабочих, служащих': 'ОП ПО ППП',
                          'Программа повышения квалификации рабочих, служащих': 'ОП ПО ПК', }
    dct_descr['Программа_аббр'] = abbr_type_program[dct_descr['Тип_программы']]

    # словарь для склоненного почему
    about_type_program = {'Повышение квалификации': 'дополнительной профессиональной программе повышения квалификации',
                          'Профессиональная переподготовка': 'дополнительной профессиональной программе профессиональной переподготовки',
                          'Программа профессиональной подготовки по профессии рабочего, должности служащего': 'основной программе',
                          'Программа переподготовки рабочих, служащих': 'ОП ПО ППП',
                          'Программа повышения квалификации рабочих, служащих': 'ОП ПО ПК', }
    dct_descr['Программа_о_чем'] = about_type_program[dct_descr['Тип_программы']]



    lst_data_df = data_df.copy() # копируем датафрейм пока он содержит только данные из листа Список
    # добавляем колонки из описания программы в датафрейм данных
    for key, value in dct_descr.items():
        data_df[key] = value
    # Конвертируем датафрейм в список словарей
    data = data_df.to_dict('records')
    dct_path = copy_folder_structure(source_folder,destination_folder) # копируем структуру папок
    for source_folder,dest_folder in dct_path.items():
        for file in os.listdir(source_folder):
            if file.endswith('.docx') and not file.startswith('~$'): # получаем только файлы docx и не временные
                # определяем тип создаваемого документа
                if 'раздельный' in file.lower():
                    # Создаем в цикле документы
                    for idx, row in enumerate(data):
                        doc = DocxTemplate(f'{source_folder}/{file}')
                        context = row
                        # print(context)
                        doc.render(context)
                        # Сохраняенм файл
                        # получаем название файла и убираем недопустимые символы < > : " /\ | ? *
                        name_file = row['ФИО']
                        name_file = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', name_file)
                        type_file = re.search(r'\b(?!Шаблон)[ЁА-Я][ёа-я]+\b',file).group()
                        if type_file:
                            name_file = f'{type_file} {name_file}'

                        # проверяем файл на наличие, если файл с таким названием уже существует то добавляем окончание
                        if name_file in used_name_file:
                            name_file = f'{name_file}_{idx}'

                        doc.save(f'{dest_folder}/{name_file[:80]}.docx')
                elif 'общий' in file.lower():
                    # Список с созданными файлами
                    files_lst = []
                    # Создаем временную папку
                    with tempfile.TemporaryDirectory() as tmpdirname:
                        print('created temporary directory', tmpdirname)
                        # Создаем и сохраняем во временную папку созданные документы Word
                        for idx, row in enumerate(data):
                            doc = DocxTemplate(f'{source_folder}/{file}')
                            context = row
                            doc.render(context)
                            # Сохраняем файл
                            # очищаем от запрещенных символов
                            name_file = row['ФИО']
                            name_file = re.sub(r'[\r\b\n\t<> :"?*|\\/]', '_', name_file)

                            doc.save(f'{tmpdirname}/{name_file[:80]}_{idx}.docx')
                            # Добавляем путь к файлу в список
                            files_lst.append(f'{tmpdirname}/{name_file[:80]}_{idx}.docx')
                        # Получаем базовый файл
                        main_doc = files_lst.pop(0)
                        # Запускаем функцию
                        combine_all_docx(main_doc, files_lst, dest_folder)
                else:
                    # генерируем текущее время
                    t = time.localtime()
                    current_time = time.strftime('%H_%M_%S', t)
                    doc = DocxTemplate(f'{source_folder}/{file}')
                    context = dict()
                    context['Итог'] = lst_data_df.to_dict('records')
                    context.update(dct_descr) # добавляем словарь с описанием программы

                    doc.render(context)
                    # Сохраняенм файл
                    # получаем название файла и убираем недопустимые символы < > : " /\ | ? *
                    name_file = file.split('.docx')[0]
                    name_file = re.sub('Шаблон ','',name_file)
                    name_file = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', name_file)

                    # проверяем файл на наличие, если файл с таким названием уже существует то добавляем окончание
                    if name_file in used_name_file:
                        name_file = f'{name_file}_{idx}'

                    doc.save(f'{dest_folder}/{name_file[:80]} {current_time}.docx')










if __name__ == '__main__':
    # main_folder_template = 'data/Шаблоны'
    main_folder_template = 'data/Шаблоны'
    main_result_folder = 'data/Результат'
    main_descr_df = pd.read_excel('data/Результат/Исходник Описание.xlsx',dtype=str)
    main_data_df = pd.read_excel('data/Результат/Исходник Список.xlsx',dtype=str)
    type_program = 'ДПО'


    generate_docs(main_descr_df,main_data_df,main_folder_template,main_result_folder,type_program)

    print('Lindy Booth')









