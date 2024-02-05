import pandas as pd
import json
from datetime import datetime
import re
import tkinter as tk
from tkinter import ttk

# Открываем файл Excel
xls = pd.ExcelFile(r'C:/Users/User/PycharmProjects/Bib25/Книги.xlsx')

# Определение максимального значения для параметра nrows
max_rows = pd.read_excel(xls, "Лист1", engine='openpyxl').shape[0]

# Чтение данных с автоматическим максимальным значением nrows
df1 = pd.read_excel(xls, "Лист1", engine='openpyxl', header=0, usecols='B', nrows=max_rows)
df2 = pd.read_excel(xls, "Лист1", engine='openpyxl', header=0, usecols='C', nrows=max_rows)
df3 = pd.read_excel(xls, "Лист1", engine='openpyxl', header=0, usecols='D', nrows=max_rows)
df4 = pd.read_excel(xls, "Лист1", engine='openpyxl', header=0, usecols='E', nrows=max_rows)
df5 = pd.read_excel(xls, "Лист1", engine='openpyxl', header=0, usecols='F', nrows=max_rows)
df6 = pd.read_excel(xls, "Лист1", engine='openpyxl', header=0, usecols='G', nrows=max_rows)
df7 = pd.read_excel(xls, "Лист1", engine='openpyxl', header=0, usecols='I', nrows=max_rows)

data = {}
data['df1'] = df1.to_dict()
data['df2'] = df2.to_dict()
data['df3'] = df3.to_dict()
data['df4'] = df4.to_dict()
data['df5'] = df5.to_dict()
data['df6'] = df6.to_dict()
data['df7'] = df7.to_dict()

def datetime_to_string(obj):
    if isinstance(obj, datetime):
        return obj.strftime("%Y-%m-%d %H:%M:%S")
    raise TypeError(f"Object of type {obj.__class__.__name__} is not JSON serializable")

with open("data_book.json", "w") as file:
    json.dump(data, file, default=datetime_to_string)


# Указываем путь к файлу JSON
json_file_path = r'C:/Users/User/PycharmProjects/Bib25/data_book.json'

# Открываем файл и загружаем JSON
with open(json_file_path, 'r') as json_file:
    data_dict = json.load(json_file)

place_book = data_dict['df1']['Unnamed: 1']
author_book = data_dict['df3']['Автор']
book = data_dict['df4']['Заглавие']
part_book = data_dict['df6']['Отдел']
age_book = data_dict['df7']['Возрастная категория']

def filter_department_84():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)

    # Пройти по всем записям в словаре и вывести только те, где значение 'Отдел' начинается с "84"
    for key in author_book.keys():
        place = place_book[key]
        author = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]

        # Проверить, является ли значение 'Отдел' строкой или целым числом, затем преобразовать его в строку
        str_values_df5 = str(values_df5) if not isinstance(values_df5, str) else values_df5

        # Извлечение значений 'Отдел', где ключ начинается с "84"
        filtered_values_df5 = [value for value in str_values_df5.split(',') if value.strip().startswith("84")]

        # Проверить, есть ли соответствующие значения 'Отдел'
        if filtered_values_df5:
            # Использовать регулярное выражение для удаления цифр слева от каждой записи
            key_without_digits = re.sub(r'^\d+:', '', key).lstrip().strip()

            # Добавить информацию в текстовое поле
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')



def filter_department_65():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)

    # Пройти по всем записям в словаре и вывести только те, где значение 'Отдел' начинается с "84"
    for key in author_book.keys():
        place = place_book[key]
        author = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]

        # Проверить, является ли значение 'Отдел' строкой или целым числом, затем преобразовать его в строку
        str_values_df5 = str(values_df5) if not isinstance(values_df5, str) else values_df5

        # Извлечение значений 'Отдел', где ключ начинается с "84"
        filtered_values_df5 = [value for value in str_values_df5.split(',') if value.strip().startswith("65")]

        # Проверить, есть ли соответствующие значения 'Отдел'
        if filtered_values_df5:
            # Использовать регулярное выражение для удаления цифр слева от каждой записи
            key_without_digits = re.sub(r'^\d+:', '', key).lstrip().strip()

            # Добавить информацию в текстовое поле
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')

def filter_department_56():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)

    # Пройти по всем записям в словаре и вывести только те, где значение 'Отдел' начинается с "84"
    for key in author_book.keys():
        place = place_book[key]
        author = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]

        # Проверить, является ли значение 'Отдел' строкой или целым числом, затем преобразовать его в строку
        str_values_df5 = str(values_df5) if not isinstance(values_df5, str) else values_df5

        # Извлечение значений 'Отдел', где ключ начинается с "84"
        filtered_values_df5 = [value for value in str_values_df5.split(',') if value.strip().startswith("56")]

        # Проверить, есть ли соответствующие значения 'Отдел'
        if filtered_values_df5:
            # Использовать регулярное выражение для удаления цифр слева от каждой записи
            key_without_digits = re.sub(r'^\d+:', '', key).lstrip().strip()

            # Добавить информацию в текстовое поле
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')

def filter_department_20():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)

    # Пройти по всем записям в словаре и вывести только те, где значение 'Отдел' начинается с "84"
    for key in author_book.keys():
        place = place_book[key]
        author = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]

        # Проверить, является ли значение 'Отдел' строкой или целым числом, затем преобразовать его в строку
        str_values_df5 = str(values_df5) if not isinstance(values_df5, str) else values_df5

        # Извлечение значений 'Отдел', где ключ начинается с "84"
        filtered_values_df5 = [value for value in str_values_df5.split(',') if value.strip().startswith("20")]

        # Проверить, есть ли соответствующие значения 'Отдел'
        if filtered_values_df5:
            # Использовать регулярное выражение для удаления цифр слева от каждой записи
            key_without_digits = re.sub(r'^\d+:', '', key).lstrip().strip()

            # Добавить информацию в текстовое поле
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')

def filter_department_22():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)

    # Пройти по всем записям в словаре и вывести только те, где значение 'Отдел' начинается с "84"
    for key in author_book.keys():
        place = place_book[key]
        author = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]

        # Проверить, является ли значение 'Отдел' строкой или целым числом, затем преобразовать его в строку
        str_values_df5 = str(values_df5) if not isinstance(values_df5, str) else values_df5

        # Извлечение значений 'Отдел', где ключ начинается с "84"
        filtered_values_df5 = [value for value in str_values_df5.split(',') if value.strip().startswith("22")]

        # Проверить, есть ли соответствующие значения 'Отдел'
        if filtered_values_df5:
            # Использовать регулярное выражение для удаления цифр слева от каждой записи
            key_without_digits = re.sub(r'^\d+:', '', key).lstrip().strip()

            # Добавить информацию в текстовое поле
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')


def filter_department_24():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)

    # Пройти по всем записям в словаре и вывести только те, где значение 'Отдел' начинается с "84"
    for key in author_book.keys():
        place = place_book[key]
        author = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]

        # Проверить, является ли значение 'Отдел' строкой или целым числом, затем преобразовать его в строку
        str_values_df5 = str(values_df5) if not isinstance(values_df5, str) else values_df5

        # Извлечение значений 'Отдел', где ключ начинается с "84"
        filtered_values_df5 = [value for value in str_values_df5.split(',') if value.strip().startswith("24")]

        # Проверить, есть ли соответствующие значения 'Отдел'
        if filtered_values_df5:
            # Использовать регулярное выражение для удаления цифр слева от каждой записи
            key_without_digits = re.sub(r'^\d+:', '', key).lstrip().strip()

            # Добавить информацию в текстовое поле
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')

def filter_department_26():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)

    # Пройти по всем записям в словаре и вывести только те, где значение 'Отдел' начинается с "84"
    for key in author_book.keys():
        place = place_book[key]
        author = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]

        # Проверить, является ли значение 'Отдел' строкой или целым числом, затем преобразовать его в строку
        str_values_df5 = str(values_df5) if not isinstance(values_df5, str) else values_df5

        # Извлечение значений 'Отдел', где ключ начинается с "84"
        filtered_values_df5 = [value for value in str_values_df5.split(',') if value.strip().startswith("26")]

        # Проверить, есть ли соответствующие значения 'Отдел'
        if filtered_values_df5:
            # Использовать регулярное выражение для удаления цифр слева от каждой записи
            key_without_digits = re.sub(r'^\d+:', '', key).lstrip().strip()

            # Добавить информацию в текстовое поле
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')



def filter_department_28():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)

    # Пройти по всем записям в словаре и вывести только те, где значение 'Отдел' начинается с "84"
    for key in author_book.keys():
        place = place_book[key]
        author = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]

        # Проверить, является ли значение 'Отдел' строкой или целым числом, затем преобразовать его в строку
        str_values_df5 = str(values_df5) if not isinstance(values_df5, str) else values_df5

        # Извлечение значений 'Отдел', где ключ начинается с "84"
        filtered_values_df5 = [value for value in str_values_df5.split(',') if value.strip().startswith("28")]

        # Проверить, есть ли соответствующие значения 'Отдел'
        if filtered_values_df5:
            # Использовать регулярное выражение для удаления цифр слева от каждой записи
            key_without_digits = re.sub(r'^\d+:', '', key).lstrip().strip()

            # Добавить информацию в текстовое поле
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')


# Your data processing code remains unchanged

def search_by_author():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)
    # Get the author from the entry field
    author_to_search = entry_author.get()



    # Perform the search logic based on the entered author
    for key in author_book:
        # Check if the value is a string before applying lower()
        place = place_book[key]
        author_value = author_book[key]
        values_df3 = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]
        if isinstance(author_value, str) and author_to_search.lower() in author_value.lower():
            # Do something with the matching author (for now, just print it)
            #print(f"Author found: {author_value}")
            text.insert(tk.END,f" {place}, Автор - {author_value}, Книга - {values_df3}, Возрастная категория - {values_df7}\n")
            print('')





def search_by_book():
    # Очистить текстовое поле перед каждым выводом
    text.delete('1.0', tk.END)
    # Get the author from the entry field
    book_to_search = entry_book.get()



    # Выполнить логику поиска на основе введенной книги
    for key in book:
        # Проверить, является ли значение строкой, прежде чем применять lower()
        place = place_book[key]
        author = author_book[key]
        book_value = book[key]
        values_df5 = part_book[key]
        values_df7 = age_book[key]
        if isinstance(book_value, str) and book_to_search.lower() in book_value.lower():
            # Выполнить действия с совпадающей книгой (в данном случае, просто вывести ее)
            text.insert(tk.END,f" {place}, Автор - {author}, Книга - {book_value}, Возрастная категория - {values_df7}\n")
            print('')




# Создать графический интерфейс Tkinter

# Create the Tkinter GUI
root = tk.Tk()

# Создаем объект класса ttk.Style
style = ttk.Style()

# Устанавливаем стиль для кнопок
style.configure("LightBlue.TButton",
                padding=1,
                foreground="#004D40",
                background="#B2DFDB",
                font="helvetica 10")

# Создание панели меню
menu_bar = tk.Menu(root)

# Создание выпадающего меню "Файл"
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Филология(языкознание, фольклористика, литературоведение). Художественная литература. Фольклор", command=filter_department_84)
file_menu.add_command(label="Естественные науки", command=filter_department_20)
file_menu.add_command(label="Физико-математические науки", command=filter_department_22)
file_menu.add_command(label="Химические науки", command=filter_department_24)
file_menu.add_command(label="Наука о Земле (геодезия, геофизика, геология и география)", command=filter_department_26)
file_menu.add_command(label="Экономика", command=filter_department_65)
file_menu.add_command(label="Биологические науки", command=filter_department_28)
file_menu.add_command(label="Клиническая медицина", command=filter_department_56)


# Добавление выпадающего меню "Файл" в панель меню
menu_bar.add_cascade(label="Жанр", menu=file_menu)

# Установка панели меню в окно
root.config(menu=menu_bar)

root.title('Список книг')
root.iconbitmap(r'icon.ico')

# Create a frame for the search entry and button at the top
search_frame = ttk.Frame(root)
search_frame.pack(side='top', fill='x')

# Entry field for searching by author
entry_author = ttk.Entry(search_frame, width=40)
entry_author.pack(side='left', padx=2, pady=2)

# Search button
button_search = ttk.Button(search_frame, text="Искать по автору", command=search_by_author, style='LightBlue.TButton')
button_search.pack(side='left', padx=2, pady=2)

# Create a frame for buttons on the left side
button_frame = ttk.Frame(root)
button_frame.pack(side='left', fill='y')

# Entry field for additional search
entry_book = ttk.Entry(search_frame, width=40)
entry_book.pack(side='left', padx=2, pady=2)

# Additional search button
button_additional = ttk.Button(search_frame, text="Искать по книге", command=search_by_book, style='LightBlue.TButton')
button_additional.pack(side='left', padx=2, pady=2)



# Create a frame for text and scrollbar on the right side
text_frame = ttk.Frame(root)
text_frame.pack(side='right', fill='both', expand=True)

# Create text widget and scrollbar inside the text frame
text = tk.Text(text_frame, height=10, width=150)
text.pack(side='left', padx=5, pady=5, fill='both', expand=True)

text_scrollbar = tk.Scrollbar(text_frame, command=text.yview)
text_scrollbar.pack(side='right', fill='y')

text.config(yscrollcommand=text_scrollbar.set)

# Start the Tkinter event loop
root.mainloop()
