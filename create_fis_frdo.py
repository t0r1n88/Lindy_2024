"""
Модуль для создания файла ФИС ФРДО
"""
import pandas as pd
import openpyxl
from tkinter import messagebox
import os

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
            df = pd.read_excel(data_file,dtype=str) # получаем данные
            lst_df = df.values.T.tolist() # превращаем в список списков по колонкам
            print(lst_df)
            template_dpo = openpyxl.load_workbook(f'{folder_template}/ФИС-ФРДО/Шаблон ФИС-ФРДО ДПО.xlsx')
            start_row = 2 # строка с которой будет начинаться записи


    except FileNotFoundError:
        messagebox.showerror('Создание документов ДПО,ПО',
                             f'В папке {folder_template}/ФИС-ФРДО не найден файл шаблона ФИС-ФРДО.\n'
                                                              f'Файлы должны иметь название - Шаблон ФИС-ФРДО ДПО и Шаблон ФИС-ФРДО ПО')

    # except NotFileTemplateDPO:
    #     messagebox.showerror('Создание документов ДПО,ПО',f'В папке {folder_template} не найден файл шаблона ФИС-ФРДО ДПО\n'
    #                                                       f'Файл должен иметь название - Шаблон ФИС-ФРДО ДПО')
    # except NotFileTemplatePO:
    #     messagebox.showerror('Создание документов ДПО,ПО',f'В папке {folder_template} не найден файл шаблона ФИС-ФРДО ПО\n'
    #                                                       f'Файл должен иметь название - Шаблон ФИС-ФРДО ПО')
















if __name__ == '__main__':
    main_data_file = 'data/Файл с данными.xlsx'
    main_folder_template = 'data/Шаблоны'
    main_result_folder = 'data/Результат'
    main_type_program = 'ДПО'

    create_fis_frdo(main_data_file,main_folder_template,main_result_folder,main_type_program)
    print('Lindy Booth !!!')

