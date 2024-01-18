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

class NotFolderSource(Exception):
    """
    Исключение для обработки случая когда не найдены папки внутри исходной папки
    """
    pass





def copy_folder_structure(source_folder:str,destination_folder:str):
    """
    Функция для копирования структуры папок внутри выбраной папки
    :param source_folder: Исходная папка
    :param destination_folder: конечная папка
    :return: Структура папок как в исходной папке
    """
    # Получаем список папок внутри source_folder
    # subfolders = [f for f in os.listdir(source_folder) if os.path.isdir(os.path.join(source_folder, f))]
    # print(subfolders)
    lst_source_folders = [] # список для хранения путей к папкам в исходной папке

    for dirname, dirnames, filenames in os.walk(source_folder):
        # print path to all subdirectories first.
        for subdirname in dirnames:
            lst_source_folders.append(f'{dirname}/{subdirname}')
    # заменяем папку назначения
    lst_dest_folders = [path.replace(source_folder,destination_folder) for path in lst_source_folders]
    for path_folder in lst_dest_folders:
        if not os.path.exists(path_folder):
            os.makedirs(path_folder)
    # создаем словарь где ключ это путь к папкам в исходном файле а значение это путь к папкам в конечной папке
    dct_path = dict(zip(lst_source_folders,lst_dest_folders))
    if len(dct_path) == 0:
        raise NotFolderSource
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
    st_multi_docs = {'удостоверение','справка','согласие','сертификат','заявление'} # список документов для которых нужно генерировать много файлов
    # Словарь для получения длинных выражений по типу программы
    where_type_program = {'Повышение квалификации':'на дополнительную профессиональную программу повышения квалификации',
                          'Профессиональная переподготовка':'на дополнительную профессиональную программу профессиональной переподготовки',
                          'Программа профессиональной подготовки по профессии рабочего, должности служащего':'на основную программу профессионального обучения профессиональной подготовки',
                          'Программа переподготовки рабочих, служащих':'на основную программу профессионального обучения профессиональной переподготовки',
                          'Программа повышения квалификации рабочих, служащих':'на основную программу профессионального обучения повышения квалификации рабочих, служащих',}
    dct_descr['Программа_куда'] = where_type_program[dct_descr['Тип_программы']]
    print(dct_descr['Программа_куда'])
    dct_path = copy_folder_structure(source_folder,destination_folder) # копируем структуру папок
    for source_folder,dest_folder in dct_path.items():
        for file in os.listdir(source_folder):
            if file.endswith('.docx') and not file.startswith('~$'): # получаем только файлы docx и не временные
                # определяем тип создаваемого документа
                print(file)
                type_doc = re.search(r'\b[Д]*ПО\b',file).group()
                print(type_doc)








if __name__ == '__main__':
    # main_folder_template = 'data/Шаблоны'
    main_folder_template = 'data/Шаблоны'
    main_result_folder = 'data/Результат'
    main_descr_df = pd.read_excel('data/Результат/Исходник Описание.xlsx',dtype=str)
    main_data_df = pd.read_excel('data/Результат/Исходник Список.xlsx',dtype=str)
    type_program = 'ДПО'


    generate_docs(main_descr_df,main_data_df,main_folder_template,main_result_folder,type_program)

    print('Lindy Booth')









