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
    print(source_folder)
    print(destination_folder)

    for dirname, dirnames, filenames in os.walk(source_folder):
        # print path to all subdirectories first.
        for subdirname in dirnames:
            lst_source_folders.append(f'{dirname}/{subdirname}')
    print(lst_source_folders)
    # заменяем папку назначения
    lst_dest_folders = [path.replace(source_folder,destination_folder) for path in lst_source_folders]
    print(lst_dest_folders)
    for path_folder in lst_dest_folders:
        if not os.path.exists(path_folder):
            os.makedirs(path_folder)
    # создаем словарь где ключ это путь к папкам в исходном файле а значение это путь к папкам в конечной папке
    dct_path = dict(zip(lst_source_folders,lst_dest_folders))
    if len(dct_path) == 0:
        raise NotFolderSource
    return dct_path


def generate_docs(descr_df:pd.DataFrame,data_df:pd.DataFrame,source_folder:str,destination_folder:str,type_program:str):
    """
    Основная функция генерации документов
    :param descr_df: датафрейм с описанием курса
    :param data_df: датафрейм с данными слушателей
    :param source_folder: исходная папка
    :param destination_folder: конечная папка
    :param type_program: тип программы ДПО или ПО
    :return: Сопроводительная документация в формате docx
    """
    dct_path = copy_folder_structure(source_folder,destination_folder) # копируем структуру папок
    for source_folder,dest_folder in dct_path.items():
        for file in os.listdir(source_folder):
            if file.endswith('.docx') and not file.startswith('~$'): # получаем только файлы docx и не временные
                print(file)








if __name__ == '__main__':
    # main_folder_template = 'data/Шаблоны'
    main_folder_template = 'data/Шаблоны'
    main_result_folder = 'data/Результат'
    main_descr_df = pd.read_excel('data/Результат/Исходник Описание.xlsx',dtype=str)
    main_data_df = pd.read_excel('data/Результат/Исходник Список.xlsx',dtype=str)
    type_program = 'ДПО'


    generate_docs(main_descr_df,main_data_df,main_folder_template,main_result_folder,type_program)

    print('Lindy Booth')









