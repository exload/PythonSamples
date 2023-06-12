"""Скрипт преобразует список из xlsx в xlsx нужного формата \
    Поиск файла осуществляется в папке со скриптом. """
import sys
import argparse
import re
import tkinter as tk
import pandas as pd


# Порядок колонок в исходном файле
READ_COLOMN_NAMES = (список колонок)
# Порядок колонок в результирующем файле
WRITE_COLOMN_NAMES = [список колонок]
# Кортеж корректных названий систем
VALID_SYSTEM_NAMES = (список систем)
# Словарь корректных значений
ACCESS_TYPE_VALID_VALUES = {'система': ['значение', 'значение'],
                            'система': ['значение', 'значение', 'значение'],
                            'система': ['значение']}
# Словарь комбинаций 
TECH_NAME_WITH_EMPTY_VALUES = {'система': ['значение'],
                               'система': ['значение'],
                               'система': ['значение']}
# Словарь комбинаций
TECH_NAME_WITH_STATIC_VALUES = {'система': ['значение']}
# Список 
NAMES_1 = ('значение', 'значение', 'значение', 'значение')
# Словарь
VALUE_1 = {'ключ': 'значение', 'ключ': 'значение',
                           'ключ': 'значение', 'ключ': 'значение'}
# RegExp выражения для проверки корректности полей
REG_1 = re.compile(r'1\d{7}')
REG_2 = re.compile(r'3\d{7}')
REG_GUID = re.compile(r'^([0-9a-fA-F]){8}-(([0-9a-fA-F]){4}-){3}([0-9a-fA-F]){12}$')
# Список ошибок обнаруженных при обработке
CONTENT_ERRORS = list()
# Количество секунд перед автоматическим закрытием окна
CLOSE_DELAY = 3
# Настройка размеров окна
WINDOW_WIDTH = 750
WINDOW_HEIGHT = 460


def set_parser(file_paths):
    """Настройка подсказок аргументов CLI-вызова скрипта"""
    parser = argparse.ArgumentParser(description="тут было описание")
    parser.add_argument('-i', metavar='input.xlsx', default='input.xlsx',
                        help='Имя файла с данными для обработки.\
                            Файл должен находится в одной папке со скриптом.')
    parser.add_argument('-o', metavar='имя_файла.xlsx',
                        default='имя_файла.xlsx',
                        help='Имя файла после обработки. Файл будет создан \
                            в той же папке, где находится скрипт.')
    args = parser.parse_args()
    file_paths['inp'] = args.i
    file_paths['out'] = args.o


def read_file(input_file_path):
    """Чтение xlsx-файла"""
    try:
        read_df = pd.read_excel(input_file_path, header=None, names=READ_COLOMN_NAMES)
    except FileNotFoundError as ex:
        message = f'Файл {input_file_path} не найден \
            ({type(ex).__name__}, {ex.args})'
        # print(message)
        CONTENT_ERRORS.append(message)
        show_log_messages(CONTENT_ERRORS)
        sys.exit(1)
    return read_df


def prepare_data_frame():
    """Подготовка DataFrame"""
    write_df = pd.DataFrame(data=None, columns=WRITE_COLOMN_NAMES)
    return write_df


def check_1(read_series):
    out = list()
    for row_num, obj in enumerate(read_series):
        # Поиск по регулярному выражению числа начинающегося с 1 и длиной из 8 цифр
        text_id = REG_1.search(str(obj))
        # text_id = re.search(r'1\d{7}', str(obj))
        if not text_id:
            out.append('Проверьте значение')
            message = f'В столбце для ___ проверьте значение в строке {row_num + 1}'
            CONTENT_ERRORS.append(message)
            # print(message)
            continue
        out.append('text ' + text_id.group(0))
    return out


def check_2(read_series):
    out = list()
    for row_num, obj in enumerate(read_series):
        # Поиск по регулярному выражению числа начинающегося с 3 и длиной из 8 цифр
        text_id = REG_2.search(str(obj))
        # text_id = re.search(r'3\d{7}', str(obj))
        if not text_id:
            out.append('Проверьте значение')
            message = f'В столбце для ____ проверьте значение в строке {row_num + 1}'
            CONTENT_ERRORS.append(message)
            # print(message)
            continue
        out.append('text ' + text_id.group(0))
    return out


def check_3(read_series):
    out = list()
    for row_num, obj in enumerate(read_series):
        # Убираем лишние пробелы
        obj = " ".join(str(obj).split())
        if obj not in VALID_SYSTEM_NAMES:
            out.append('Проверьте значение')
            message = f'В столбце для ___ проверьте значение в строке {row_num + 1}'
            CONTENT_ERRORS.append(message)
            # print(message)
            continue
        out.append(obj)
    return out


def check_4(read_series, write_series_system_name):
    out = list()
    for row_num, obj in enumerate(read_series):
        system_name = write_series_system_name[row_num]
        if system_name not in ACCESS_TYPE_VALID_VALUES:
            out.append('Проверьте значение')
            message = f'Для проверки корректности значения для ___ проверьте значение в столбце для ____ в строке {row_num + 1}'
            CONTENT_ERRORS.append(message)
            # print(message)
            continue
        else:
            # Убираем лишние пробелы
            obj = " ".join(str(obj).split())
            if obj == 'text':
                obj = 'text'
            if obj not in ACCESS_TYPE_VALID_VALUES[system_name]:
                out.append('Проверьте значение')
                message = f'В столбце для Тип объекта проверьте значение в строке {row_num + 1}'
                CONTENT_ERRORS.append(message)
                # print(message)
                continue
            out.append(obj)
    return out


def check_5(guid):
    if REG_GUID.match(guid):
        return True
    else:
        return False


def check_6(read_series, write_series_system_name, write_series_access_type):
    out = list()
    for row_num, obj in enumerate(read_series):
        system_name = write_series_system_name[row_num]
        access_type = write_series_access_type[row_num]
        if system_name in TECH_NAME_WITH_EMPTY_VALUES and access_type in TECH_NAME_WITH_EMPTY_VALUES[system_name]:
            out.append('')
        elif system_name in TECH_NAME_WITH_STATIC_VALUES and access_type in TECH_NAME_WITH_STATIC_VALUES[system_name]:
            out.append('text')
        elif system_name == 'text' and access_type == 'text':
            # Убираем лишние пробелы
            obj = " ".join(str(obj).split())
            if not check_5(obj):
                out.append('Проверьте значение')
                message = f'В столбце для ______ проверьте значение в строке {row_num + 1}'
                CONTENT_ERRORS.append(message)
                # print(message)
                continue
            out.append(obj)
        elif system_name == 'Проверьте значение' or access_type == 'Проверьте значение':
            out.append('Проверьте значение')
            message = f'Проверьте значения в предыдущих столбцах(____ и _____) в строке {row_num + 1}'
            CONTENT_ERRORS.append(message)
            # print(message)
            continue
        elif system_name == 'text' and access_type == 'text':
            # Убираем лишние пробелы
            obj = " ".join(str(obj).split())
            if obj not in NAMES_1:
                out.append('Проверьте значение')
                message = f'В столбце для ______ проверьте значение в строке {row_num + 1}'
                CONTENT_ERRORS.append(message)
                # print(message)
                continue
            out.append(obj)
    return out


def check_7(read_series, write_series_access_type, write_series_tech_name):
    out = list()
    for row_num, obj in enumerate(read_series):
        if write_series_access_type[row_num] == 'text':
            out.append('text')
        elif write_series_tech_name[row_num] in VALUE_1:
            out.append((obj.replace(' ', '')).upper())
        else:
            out.append('')
    return out


def fill_output_dataframe(read_df, write_df):
    """Формирование DataFrame для результирующего файла"""
    write_df['text'] = pd.Series(check_1(read_df['text_id']))
    write_df['text'] = read_df['text'].copy()
    write_df['text'] = 'Нет'
    write_df['text'] = pd.Series(check_2(read_df['text_id']))
    write_df['text'] = read_df['text'].copy()
    write_df['text'] = pd.Series(check_3(read_df['text']))
    write_df['text'] = pd.Series(check_4(read_df['text'], write_df['text']))
    write_df['text'] = pd.Series(check_6(read_df['tech_name'], write_df['text'], write_df['text']))
    write_df['text'] = pd.Series(check_7(read_df['tech_value'], write_df['text'], write_df['text']))
    write_df['text'] = 'text'
    write_df['text'] = read_df['access_func_name'].copy()
    write_df['text'] = read_df['access_description'].copy()


def write_file(output_file_path, write_df):
    """Запись xlsx-файла"""
    my_sheet_name = 'text'
    try:
        write_df.to_excel(output_file_path, sheet_name=my_sheet_name,
                          header=WRITE_COLOMN_NAMES, index=False)
    except Exception as ex:
        message = f'Неудалось создать файл {output_file_path}. Проверьте права. ({type(ex).__name__}, {ex.args})'
        # print(message)
        CONTENT_ERRORS.append(message)
        show_log_messages(CONTENT_ERRORS)
        sys.exit(1)


def show_log_messages(messages):
    """Отображение окна с ошибками"""
    errors_num = len(messages)
    if errors_num != 0:
        messages = '\n'.join(messages)
        no_errors_flag = False
    else:
        messages = 'Ошибок не обнаружено.'
        no_errors_flag = True
    root = tk.Tk()
    root.title('Окно вывода ошибок')
    # Настройка позиционирования окна
    screen_width = root.winfo_screenwidth()
    screen_heigth = root.winfo_screenheight()
    pos_x = int(abs((screen_width / 2) - (WINDOW_WIDTH / 2)))
    pos_y = int(abs((screen_heigth / 2) - (WINDOW_HEIGHT / 2)))
    root.geometry(f'{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{pos_x}+{pos_y}')
    # Запрет изменения размера окна
    root.resizable(False, False)
    # Создание фреймов для позиционирования виджетов
    frame_top = tk.Frame(root)
    frame_bottom = tk.Frame(root)
    # Полоса прокрутки
    scroll = tk.Scrollbar(frame_top, orient='vertical')
    scroll.pack(side='right', fill='y')
    # Текстовое поле
    # yscrollcommand - привязка к текстовому виджету полосы прокрутки
    txt = tk.Text(frame_top, wrap='word', yscrollcommand=scroll.set, width=WINDOW_WIDTH)
    txt.insert(1.0, messages)
    # привязка полосы прокрутки к текстовому виджету
    scroll.config(command=txt.yview)
    # Кнопка
    btn = tk.Button(frame_bottom, text='Закрыть', command=root.destroy, height=2, width=20)
    # Текстовое поле
    if no_errors_flag:
        lbl = tk.Label(frame_bottom, text=f'Окно автоматически закроется через {CLOSE_DELAY}сек.')
        lbl.pack()
    # Отображение виджетов и фреймов
    txt.pack()
    btn.pack()
    frame_top.pack()
    frame_bottom.pack()
    if no_errors_flag:
        # Вызов функции закрытия окна через DELAY_CLOSE секунд
        root.after(CLOSE_DELAY * 1000, root.destroy)
    # Отображение окна
    root.mainloop()


def main():
    """Основная функция"""
    file_paths = dict()
    read_df = pd.DataFrame()
    write_df = pd.DataFrame()
    set_parser(file_paths)
    read_df = read_file(file_paths['inp'])
    write_df = prepare_data_frame()
    fill_output_dataframe(read_df, write_df)
    write_file(file_paths['out'], write_df)
    show_log_messages(CONTENT_ERRORS)


if __name__ == '__main__':
    main()
