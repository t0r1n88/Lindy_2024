from create_documents import create_docs # импортируем основную функцию генерации документов
from create_doc_legal_person import create_docs_legal_person # импортируем функцию для генерации документов юрлиц
from prepare_data import prepare_list # импортируем функцию очистки данных
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import datetime
import os
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
import sys
import locale
import logging
logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)
"""
Системные функции для подгрузки логотипа и работы полосы прокрутки
"""

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


"""
Функции для создания контекстного меню(Копировать,вставить,вырезать)
"""


def make_textmenu(root):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # эта штука делает меню
    global the_menu
    the_menu = Menu(root, tearoff=0)
    the_menu.add_command(label="Вырезать")
    the_menu.add_command(label="Копировать")
    the_menu.add_command(label="Вставить")
    the_menu.add_separator()
    the_menu.add_command(label="Выбрать все")


def callback_select_all(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # select text after 50ms
    window.after(50, lambda: event.widget.select_range(0, 'end'))


def show_textmenu(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    e_widget = event.widget
    the_menu.entryconfigure("Вырезать", command=lambda: e_widget.event_generate("<<Cut>>"))
    the_menu.entryconfigure("Копировать", command=lambda: e_widget.event_generate("<<Copy>>"))
    the_menu.entryconfigure("Вставить", command=lambda: e_widget.event_generate("<<Paste>>"))
    the_menu.entryconfigure("Выбрать все", command=lambda: e_widget.select_range(0, 'end'))
    the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)



def on_scroll(*args):
    canvas.yview(*args)

def set_window_size(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Устанавливаем размер окна в 80% от ширины и высоты экрана
    if screen_width >= 3840:
        width = int(screen_width * 0.2)
    elif screen_width >= 2560:
        width = int(screen_width * 0.31)
    elif screen_width >= 1920:
        width = int(screen_width * 0.41)
    elif screen_width >= 1600:
        width = int(screen_width * 0.5)
    elif screen_width >= 1280:
        width = int(screen_width * 0.62)
    elif screen_width >= 1024:
        width = int(screen_width * 0.77)
    else:
        width = int(screen_width * 1)

    height = int(screen_height * 0.8)

    # Рассчитываем координаты для центрирования окна
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размер и положение окна
    window.geometry(f"{width}x{height}+{x}+{y}")

"""
Функции для получения путей к папкам и файлам
"""

"""
Функции для выбора папок и файлов при генерации документов
"""

def select_file_data_create_docs():
    """
    Функция для выбора файла xlsx с данными программы
    :return:
    """
    global data_create_docs
    # Получаем путь к файлу
    data_create_docs = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_template_folder():
    """
    Функия для выбора папки с шаблонами документов
    :return:
    """
    global template_folder
    template_folder = filedialog.askdirectory()

def select_result_folder():
    """
    Функия для выбора папки с шаблонами документов
    :return:
    """
    global result_folder
    result_folder = filedialog.askdirectory()


def processing_create_docs():
    """
    Функция для запуска создания документов
    :return: Документы в конечной папке
    """
    try:
        create_docs(data_create_docs,template_folder,result_folder)
    except NameError:
        messagebox.showerror('Создание документов ДПО,ПО',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')


"""
Функции для выбора папок и файлов при генерации документов юрлиц
"""

def select_file_data_create_docs_legal_person():
    """
    Функция для выбора файла xlsx с данными программы
    :return:
    """
    global data_create_docs_legal_person
    # Получаем путь к файлу
    data_create_docs_legal_person = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_template_folder_legal_person():
    """
    Функия для выбора папки с шаблонами документов
    :return:
    """
    global template_folder_legal_person
    template_folder_legal_person = filedialog.askdirectory()

def select_result_folder_legal_person():
    """
    Функия для выбора папки с шаблонами документов
    :return:
    """
    global result_folder_legal_person
    result_folder_legal_person = filedialog.askdirectory()


def processing_create_docs_legal_person():
    """
    Функция для запуска создания документов
    :return: Документы в конечной папке
    """
    try:
        create_docs_legal_person(data_create_docs_legal_person,template_folder_legal_person,result_folder_legal_person)
    except NameError:
        messagebox.showerror('Создание документов ДПО,ПО',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')


"""
Функции для вкладки подготовка файлов
"""
def select_prep_file():
    """
    Функция для выбора файла который нужно преобразовать
    :return:
    """
    global glob_prep_file
    # Получаем путь к файлу
    glob_prep_file = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_prep():
    """
    Функция для выбора папки куда будет сохранен преобразованный файл
    :return:
    """
    global glob_path_to_end_folder_prep
    glob_path_to_end_folder_prep = filedialog.askdirectory()


def processing_preparation_file():
    """
    Функция для генерации документов
    """
    try:
        prepare_list(glob_prep_file,glob_path_to_end_folder_prep)

    except NameError:
        messagebox.showerror('Линди Создание документов ДПО,ПО',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')




if __name__ == '__main__':
    window = Tk()
    window.title('Линди Создание документов ДПО,ПО ver 2.1')
    # Устанавливаем размер и положение окна
    set_window_size(window)

    window.resizable(True, True)
    # Добавляем контекстное меню в поля ввода
    make_textmenu(window)

    # Создаем вертикальный скроллбар
    scrollbar = Scrollbar(window, orient="vertical")

    # Создаем холст
    canvas = Canvas(window, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)

    # Привязываем скроллбар к холсту
    scrollbar.config(command=canvas.yview)

    # Создаем ноутбук (вкладки)
    tab_control = ttk.Notebook(canvas)

    tab_create_docs = ttk.Frame(tab_control)
    tab_control.add(tab_create_docs, text='Создание документов ДПО и ПО')

    create_docs_frame_description = LabelFrame(tab_create_docs)
    create_docs_frame_description.pack()

    lbl_hello_create_docs = Label(create_docs_frame_description,
                                  text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                       'Создание сопроводительной документации к программам ДПО и ПО\n',
                                  width=60)
    lbl_hello_create_docs.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_docs = resource_path('logo.png')
    img_create_docs = PhotoImage(file=path_to_img_create_docs)
    Label(create_docs_frame_description,
          image=img_create_docs, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_create_docs = LabelFrame(tab_create_docs, text='Подготовка')
    frame_data_create_docs.pack(padx=10, pady=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_create_docs_file= Button(frame_data_create_docs, text='1) Выберите файл', font=('Arial Bold', 14),
                                       command=select_file_data_create_docs)
    btn_choose_create_docs_file.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_template_folder = Button(frame_data_create_docs, text='2) Выберите папку с шаблонами', font=('Arial Bold', 14),
                                        command=select_template_folder)
    btn_choose_template_folder.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_result_folder = Button(frame_data_create_docs, text='3) Выберите конечную папку', font=('Arial Bold', 14),
                                        command=select_result_folder)
    btn_choose_result_folder.pack(padx=10, pady=10)

    # Создаем кнопку генерации документов
    btn_process_create_docs = Button(tab_create_docs,text='4) Создать документы', font=('Arial Bold', 14),
                                        command=processing_create_docs)
    btn_process_create_docs.pack(padx=10, pady=10)

    """
    Создаем вкладку для генерации документов юрлиц
    """

    tab_create_docs_legal_person = ttk.Frame(tab_control)
    tab_control.add(tab_create_docs_legal_person, text='Создание документов ДПО и ПО для юрлиц')

    create_docs_legal_person_frame_description = LabelFrame(tab_create_docs_legal_person)
    create_docs_legal_person_frame_description.pack()

    lbl_hello_create_docs_legal_person = Label(create_docs_legal_person_frame_description,
                                               text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                                    'Создание сопроводительной документации к программам ДПО и ПО для юрлиц\n',
                                               width=60)
    lbl_hello_create_docs_legal_person.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_docs_legal_person = resource_path('logo.png')
    img_create_docs_legal_person = PhotoImage(file=path_to_img_create_docs_legal_person)
    Label(create_docs_legal_person_frame_description,
          image=img_create_docs_legal_person, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_create_docs_legal_person = LabelFrame(tab_create_docs_legal_person, text='Подготовка')
    frame_data_create_docs_legal_person.pack(padx=10, pady=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_create_docs_legal_person_file = Button(frame_data_create_docs_legal_person, text='1) Выберите файл',
                                                      font=('Arial Bold', 14),
                                                      command=select_file_data_create_docs_legal_person)
    btn_choose_create_docs_legal_person_file.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_template_folder_legal_person = Button(frame_data_create_docs_legal_person, text='2) Выберите папку с шаблонами',
                                        font=('Arial Bold', 14),
                                        command=select_template_folder_legal_person)
    btn_choose_template_folder_legal_person.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_result_folder_legal_person = Button(frame_data_create_docs_legal_person, text='3) Выберите конечную папку',
                                      font=('Arial Bold', 14),
                                      command=select_result_folder_legal_person)
    btn_choose_result_folder_legal_person.pack(padx=10, pady=10)

    # Создаем кнопку генерации документов
    btn_process_create_docs_legal_person = Button(tab_create_docs_legal_person, text='4) Создать документы',
                                                  font=('Arial Bold', 14),
                                                  command=processing_create_docs_legal_person)
    btn_process_create_docs_legal_person.pack(padx=10, pady=10)

    """
       Создаем вкладку для предварительной обработки списка
       """
    tab_preparation = ttk.Frame(tab_control)
    tab_control.add(tab_preparation, text='Обработка списка')

    preparation_frame_description = LabelFrame(tab_preparation)
    preparation_frame_description.pack()

    lbl_hello_preparation = Label(preparation_frame_description,
                                  text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                       'Очистка от лишних пробелов и символов; поиск пропущенных значений\n в колонках с персональными данными,'
                                       '(ФИО,паспортные данные,\nтелефон,e-mail,дата рождения,ИНН)\n преобразование СНИЛС в формат ХХХ-ХХХ-ХХХ ХХ.\n'
                                       'Создание списка дубликатов по каждой колонке\n'
                                       'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                                       'Для корректной работы программы уберите из таблицы\nобъединенные ячейки',
                                  width=60)
    lbl_hello_preparation.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_preparation = resource_path('logo.png')
    img_preparation = PhotoImage(file=path_to_img_preparation)
    Label(preparation_frame_description,
          image=img_preparation, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_prep = LabelFrame(tab_preparation, text='Подготовка')
    frame_data_prep.pack(padx=10, pady=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_prep_file = Button(frame_data_prep, text='1) Выберите файл', font=('Arial Bold', 14),
                                  command=select_prep_file)
    btn_choose_prep_file.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_prep = Button(frame_data_prep, text='2) Выберите конечную папку', font=('Arial Bold', 14),
                                        command=select_end_folder_prep)
    btn_choose_end_folder_prep.pack(padx=10, pady=10)

    # Создаем кнопку очистки
    btn_choose_processing_prep = Button(tab_preparation, text='3) Выполнить обработку', font=('Arial Bold', 20),
                                        command=processing_preparation_file)
    btn_choose_processing_prep.pack(padx=10, pady=10)









    














    # Создаем виджет для управления полосой прокрутки
    canvas.create_window((0, 0), window=tab_control, anchor="nw")

    # Конфигурируем холст для обработки скроллинга
    canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))
    scrollbar.pack(side="right", fill="y")

    # Вешаем событие скроллинга
    canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)
    window.mainloop()