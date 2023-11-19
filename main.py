import subprocess
import os

import imageio as imageio
import openai
import openpyxl

import win32com.client.dynamic
import pyttsx3
import speech_recognition as sr

import time

from pywinauto.application import Application
import threading

import tkinter as tk
from PIL import Image, ImageTk

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

filename = 'РАСПИСАНИЕ.xlsx'
current_directory = os.getcwd()
file_path = os.path.join(current_directory, filename)


def open_excel_file(file_path):
    try:
        subprocess.Popen(['start', 'excel', '/e', '/x', file_path], shell=True)
        return True
    except Exception as e:
        print(f"Ошибка при открытии файла Excel: {e}")
        return False


def voice(text):
    engine.say(text)
    engine.runAndWait()


def find_cell_address(file_path, search_words):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for row in sheet.iter_rows():
        for cell in row:
            if cell.value in search_words:
                return cell

    return None


def save_menu(text):
    current_directory = os.getcwd()
    filepath = os.path.join(current_directory, "menu.txt")

    try:
        with open(filepath, "w", encoding="utf-8") as file:
            file.write(text)
        print("Файл успешно сохранен.")
    except Exception as e:
        print(f"Ошибка сохранения файла: {e}")


# Сохранение событий
def save_events(events_text):
    current_directory = os.getcwd()
    filepath = os.path.join(current_directory, "events.txt")

    try:
        with open(filepath, "w", encoding="utf-8") as file:
            file.write(events_text)
        print("Файл успешно сохранен.")
    except Exception as e:
        print(f"Ошибка сохранения файла: {e}")


# Сохранение длительности звонков
def save_call_duration(call_duration_text):
    current_directory = os.getcwd()
    filepath = os.path.join(current_directory, "call_duration.txt")

    try:
        with open(filepath, "w", encoding="utf-8") as file:
            file.write(call_duration_text)
        print("Файл успешно сохранен.")
    except Exception as e:
        print(f"Ошибка сохранения файла: {e}")


def maximize_excel_window():
    try:
        app = Application(backend="uia").connect(title_re='.*Excel', visible_only=True)
        window = app.window(title_re='.*Excel')
        window.restore()
        window.maximize()
        return True
    except Exception as e:
        print(f"Ошибка при максимизации окна Excel: {e}")
        return False


def navigate_to_cell(cell):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    # Открываем файл
    workbook = excel.Workbooks.Open(file_path)

    # Активируем нужный лист
    worksheet = workbook.Worksheets(1)
    worksheet.Activate()

    # Переходим к ячейке
    cell_range = worksheet.Range(cell.coordinate)
    cell_range.Select()

    # Минимизируем и разворачиваем окно Excel

    excel.WindowState = 1
    worksheet.Application.ActiveWindow.ScrollRow += 8

    time.sleep(30)

    # Отображаем главное окно программы
    root.deiconify()


def read_menu():
    try:
        with open('menu.txt', 'r', encoding='utf-8') as file:
            menu = file.read()
            return menu
    except Exception as e:
        print(f"Ошибка при чтении файла меню: {e}")
        return None


# Чтение событий из файла
def read_events():
    try:
        with open('events.txt', 'r', encoding='utf-8') as file:
            events_text = file.read()
            return events_text
    except Exception as e:
        print(f"Ошибка при чтении файла событий: {e}")
        return None


# Чтение длительности звонков из файла
def read_call_duration():
    try:
        with open('call_duration.txt', 'r', encoding='utf-8') as file:
            call_duration_text = file.read()
            return call_duration_text
    except Exception as e:
        print(f"Ошибка при чтении файла длительности звонков: {e}")
        return None


def save_news(news_text):
    current_directory = os.getcwd()
    filepath = os.path.join(current_directory, "news.txt")

    try:
        with open(filepath, "w", encoding="utf-8") as file:
            file.write(news_text)
        print("Файл успешно сохранен.")
    except Exception as e:
        print(f"Ошибка сохранения файла: {e}")


# Чтение новостей из файла
def read_news():
    try:
        with open('news.txt', 'r', encoding='utf-8') as file:
            news_text = file.read()
            return news_text
    except Exception as e:
        print(f"Ошибка при чтении новостей: {e}")
        return None


def open_settings():
    def check_password():
        # Функция для проверки пароля
        password = entry_password.get()
        if password == "gym19":
            password_dialog.destroy()
            display_settings()
        else:
            messagebox.showerror("Ошибка", "Неправильный пароль!")
            entry_password.delete(0, END)  # Очищаем поле ввода пароля

    def display_settings():
        # Функция для отображения настроек
        root.withdraw()
        settings_window = Toplevel(root)
        settings_window.title("Настройки")
        settings_window.geometry("700x700")

        tab_control = ttk.Notebook(settings_window)

        # Вкладка "Меню"

        menu_tab = ttk.Frame(tab_control)
        tab_control.add(menu_tab, text="Меню")

        menu_label = Label(menu_tab, text="Изменить меню", font=("Arial", 16, "bold"))
        menu_label.pack(pady=20)

        menu_textbox = Text(menu_tab, width=50, height=10)
        menu_textbox.pack()

        save_menu_button = Button(menu_tab, text="Сохранить меню",
                                  command=lambda: save_menu(menu_textbox.get("1.0", "end-1c")))
        save_menu_button.pack(pady=20)

        # Вкладка "Расписание звонков"
        schedule_tab = ttk.Frame(tab_control)
        tab_control.add(schedule_tab, text="Расписание звонков")

        schedule_label = Label(schedule_tab, text="Изменить расписание звонков", font=("Arial", 16, "bold"))
        schedule_label.pack(pady=20)

        call_duration_textbox = Text(schedule_tab, width=50, height=10)
        call_duration_textbox.pack()

        save_schedule_button = Button(schedule_tab, text="Сохранить расписание звонков",
                                      command=lambda: save_call_duration(call_duration_textbox.get("1.0", "end-1c")))
        save_schedule_button.pack(pady=20)

        # Вкладка "Новости"
        news_tab = ttk.Frame(tab_control)
        tab_control.add(news_tab, text="Новости")

        news_label = Label(news_tab, text="Изменить новости", font=("Arial", 16, "bold"))
        news_label.pack(pady=20)

        news_textbox = Text(news_tab, width=50, height=10)
        news_textbox.pack()

        save_news_button = Button(news_tab, text="Сохранить новости",
                                  command=lambda: save_news(news_textbox.get("1.0", "end-1c")))
        save_news_button.pack(pady=20)

        # Вкладка "События"
        events_tab = ttk.Frame(tab_control)
        tab_control.add(events_tab, text="События")

        events_label = Label(events_tab, text="Изменить события", font=("Arial", 16, "bold"))
        events_label.pack(pady=20)

        events_textbox = Text(events_tab, width=50, height=10)
        events_textbox.pack()

        save_events_button = Button(events_tab, text="Сохранить события",
                                    command=lambda: save_events(events_textbox.get("1.0", "end-1c")))
        save_events_button.pack(pady=20)

        # Размещение вкладок на окне настроек
        tab_control.pack(expand=1, fill="both")
        settings_window.protocol("WM_DELETE_WINDOW", lambda: close_settings(
            settings_window))  # Установка действия при закрытии окна настроек

        settings_window.wait_window()  # Ожидание закрытия окна настроек и возврат к главному окну

        root.deiconify()  # Показать главное окно

    def close_settings(window):
        window.destroy()
        root.deiconify()  # Показать главное окно

        # Создаем окно ввода пароля

    password_dialog = Toplevel(root)
    password_dialog.title("Ввод пароля")
    password_dialog.geometry("300x100")
    password_dialog.resizable(False, False)

    label_password = Label(password_dialog, text="Пароль:")
    label_password.pack()

    entry_password = Entry(password_dialog, show="*")
    entry_password.pack()

    btn_check_password = Button(password_dialog, text="Проверить пароль", command=check_password)
    btn_check_password.pack()


def close_settings(settings_window):
    settings_window.destroy()
    root.deiconify()  # Восстановить главное окно "Девятнашки"


def close_program():
    password = "gym19"

    def validate_password():
        entered_password = password_entry.get()

        if entered_password == password:
            password_window.destroy()
            root.destroy()
        else:
            messagebox.showerror("Ошибка", "Неверный пароль!")

    password_window = Toplevel(root)
    password_window.title("Введите пароль")
    password_window.geometry("300x100")
    password_window.resizable(False, False)

    password_label = Label(password_window, text="Пароль:")
    password_label.pack()

    password_entry = Entry(password_window, show="*")
    password_entry.pack()

    password_button = Button(password_window, text="Выйти", command=validate_password)
    password_button.pack()


last_command = ""
english_to_russian = {
    'A': 'а', 'B': 'б', 'C': 'ц', 'D': 'д', 'E': 'е', 'F': 'ф', 'G': 'г',
    'H': 'х', 'I': 'и', 'J': 'й', 'K': 'к', 'L': 'л', 'M': 'м', 'N': 'н',
    'O': 'о', 'P': 'п', 'Q': 'к', 'R': 'р', 'S': 'с', 'T': 'т', 'U': 'у',
    'V': 'в', 'W': 'в', 'X': 'кс', 'Y': 'й', 'Z': 'з'
}

# Словарь для преобразования русских числительных в цифры
russian_to_number = {
    'первый': 1, 'второй': 2, 'третий': 3, 'четвертый': 4, 'пятый': 5,
    'шестой': 6, 'седьмой': 7, 'восемой': 8, 'девятый': 9, 'десятый': 10,
    'одиннадцатый': 11,
    # Добавьте остальные числа по аналогии
}


# Функция для преобразования заглавных английских букв в строчные русские буквы
def convert_to_russian(text):
    result = ""
    is_previous_letter_cyrillic = False
    for char in text:
        if char.isalpha() and char.upper() in english_to_russian:
            result += english_to_russian[char.upper()]
            is_previous_letter_cyrillic = True
        elif char.isalpha():
            if is_previous_letter_cyrillic:
                result += char.lower()
            else:
                result += char
            is_previous_letter_cyrillic = False
        else:
            result += char
            is_previous_letter_cyrillic = False
    return result


# Функция для преобразования русских числительных в цифры
def convert_to_number_or_word(word):
    # Добавим преобразования для числительных "первыйв", "второйв" и т.д.
    conversion_dict = {
        "первыйв": "1в",
        "первыйа": "1а",
        "первыйб": "1б",
        "первыйд": "1д",
        "первыйг": "1г",
        "второйа": "2а",
        "второйб": "2б",
        "второйв": "2в",
        "второйг": "2г",
        "второйд": "2д",
        "третийа": "3а",
        "третийб": "3б",
        "третийв": "3в",
        "третийг": "3г",
        "третийд": "3д",
        "четвёртыйа": "4а",
        "четвёртыйб": "4б",
        "четвёртыйв": "4в",
        "четвёртыйг": "4г",
        "четвёртыйи": "4и",
        "4ив": "4и",
        "пятыйа": "5а",
        "пятыйб": "5б",
        "пятыйв": "5в",
        "пятыйг": "5г",
        "пятыйи": "5и",
        "пятыйматем": "5 матем",
        "5м-1": "5м1",
        "пятыйм1": "5м1",
        "пятыймодин": "5м1",
        "5м-2": "5м2",
        "пятыйм2": "5м2",
        "пятыймдва": "5м2",
        "шестойа": "6а",
        "стойа": "6а",
        "шестойб": "6б",
        "шестойв": "6в",
        "шестойл": "6л",
        "шестойматем": "6 матем",
        "6матем": "6 матем",
        "стойматем": "6 матем",
        "6м-1": "6м1",
        "шестойм1": "6м1",
        "шестоймодин": "6м1",
        "6м-2": "6м2",
        "шестойм2": "6м2",
        "шестоймдва": "6м2",
        "седьмойа": "7а",
        "седьмойб": "7б",
        "седьмойи": "7и",
        "седьмойив": "7и",
        "седьмойматем": "7 матем",
        "седьмойлодин": "7л1",
        "седьмойл1": "7л1",
        "седьмойлдва": "7л2",
        "седьмойл2": "7л2",
        "7матем": "7 матем",
        "восьмойа": "8а",
        "восьмойб": "8б",
        "восьмойи": "8и",
        "восьмойматем": "8 матем.",
        "8матем": "8 матем.",
        "восьмойлодин": "8л1",
        "восьмойл1": "8л1",
        "восьмойлдва": "8л2",
        "восьмойл2": "8л2",
        "девятыйа": "9а",
        "9 А": "9а",
        "девятыйб": "9б",
        "девятыйв": "9в",
        "девятыйматем": "9 матем",
        "9матем": "9 матем",
        "девятыйлодин": "9л1",
        "девятыйл1": "9л1",
        "девятыйлдва": "9л2",
        "девятыйл2": "9л2",
        "десятыйа": "10а",
        "десятыйл": "10л",
        "десятыйм": "10м",
        "десятыйэм": "10м",
        "десятыйам": "10м",
        "10ам": "10м",
        "10эм": "10м",
        "десятыйматем": "10 матем",
        "10матем": "10 матем",
        "одиннадцатыйа": "11а",
        "одиннадцатыйл": "11л",
        "одиннадцатыйм": "11м",
        "одиннадцатыйам": "11м",
        "одиннадцатыйэм": "11м",
        "11ам": "11м",
        "11эм": "11м",

    }

    if word in russian_to_number:
        return str(russian_to_number[word])
    elif word in conversion_dict:
        return conversion_dict[word]
    else:
        return word.lower()  # Преобразование русских заглавных букв в строчные


def process_program():
    def process_commands():
        global btn_start
        while True:
            r = sr.Recognizer()
            with sr.Microphone() as source:
                print("Скажите что-нибудь:")
                audio = r.listen(source)
                try:
                    command = r.recognize_google(audio, language="ru-RU")
                    last_command = command
                    print("Вы сказали:", command)
                except sr.UnknownValueError:
                    print("Извините, не удалось распознать речь.")
                    recreate_start_button()
                    break

            if "открыть расписание" in command.lower() or "Открыть расписание" in command.lower():
                print("Скажите номер и букву класса:")
                command_label.config(text="Скажите номер и букву класса:")
                voice("Пожалуйста, скажите номер и букву класса.")
                r = sr.Recognizer()
                with sr.Microphone() as source:
                    audio = r.listen(source)
                try:
                    search_words = r.recognize_google(audio, language="ru-RU").replace(" ", "").split(',')
                    search_words = [convert_to_russian(word) for word in search_words]

                    print("Вы сказали:", search_words)

                    command_label.config(text="Вы сказали: " + last_command)
                except sr.UnknownValueError:
                    response = "Извините, не удалось распознать ключевые слова"
                    print("Извините, не удалось распознать ключевые слова.")

                    continue
                search_words = [f"Класс - {convert_to_number_or_word(word)}" for word in search_words]

                cell = find_cell_address(file_path, search_words)
                if cell is not None:
                    root.withdraw()
                    maximize_excel_window()
                    navigate_to_cell(cell)
                    root.deiconify()

                    response = f"Расписание ( {convert_to_number_or_word(search_words[0])}) было успешно развернуто и перемещено к ячейке"
                else:
                    response = f"  ({convert_to_number_or_word(search_words[0])}) не были найдены в расписании. Пожалуйста, повторите запрос."

            elif "новости" in command.lower() or "Новости гимназии" in command.lower() or "новости гимназии" in command.lower() or "Новости" in command.lower() or "Новости в гимназии" in command.lower() or "новости в гимназии" in command.lower():
                response = read_news()
                command_label.config(text="Вы сказали: " + last_command)
            elif "меню" in command.lower() or "Меню" in command.lower():
                response = read_menu()
                command_label.config(text="Вы сказали: " + last_command)
            elif "события" in command.lower() or "события в гимназие" in command.lower():
                response = read_events()
                command_label.config(text="Вы сказали: " + last_command)
            elif "расписание звонков" in command.lower() or "Расписание звонков" in command.lower():
                response = read_call_duration()
                command_label.config(text="Вы сказали: " + last_command)
            else:
                openai.api_key = "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"
                response = "извините, повторите вопрос "

            response = response.replace("Класс - ", "")

            print("ИИ:", response)
            voice(response)
            command_label.config(text="Вы сказали: " + last_command)

            text_var.set("ИИ: " + response)
            recreate_start_button()

            break  # Выход из бесконечного цикла после обработки команды

    def recreate_start_button():
        global btn_start
        btn_start.destroy()  # Удаление предыдущей кнопки "Старт"
        btn_start = tk.Button(root, image=photo, bd=0, command=process_program)
        btn_start.config(bg='black', activebackground='black', state='normal')
        btn_start.place(relx=0.5, rely=0.10, anchor="center")
        video_label.lower()

    btn_start.destroy()
    threading.Thread(target=process_commands).start()


root = Tk()
root.attributes('-fullscreen', True)  # Открыть окно на весь экран
root['bg'] = '#fafafa'
root.title('Девятнашка')
video_path = "фон.mp4"

video = imageio.get_reader(video_path)
num_frames = video.get_length()
frame_index = 0

video_label = tk.Label(root)
video_label.pack(fill=tk.BOTH, expand=True)


def update_video():
    global frame_index
    frame_index += 1
    if frame_index >= num_frames:
        frame_index = 0

    try:
        # Извлечение текущего кадра из видео и преобразование в изображение
        frame = video.get_data(frame_index)
        image = Image.fromarray(frame)
        photo = ImageTk.PhotoImage(image)

        video_label.configure(image=photo)
        video_label.image = photo

    except IndexError:

        frame_index = 0

    video_label.after(1, update_video)


# Запуск обновления видео
update_video()

text_var = StringVar()
text_label = tk.Label(root, textvariable=text_var, font=('Arial', 18), fg='white', bg='black')
text_label.pack(anchor='w', padx=10, pady=10)

text_var.set('Запустите девятнашку')

image_path = 'старт.jpg'
image = Image.open(image_path)
image = image.resize((200, 200))  # Установка размера изображения
photo = ImageTk.PhotoImage(image)

# Создание кнопки с использованием изображения
btn_start = tk.Button(root, image=photo, bd=0, command=process_program)
btn_start.config(bg='black', activebackground='black', state='normal')
btn_start.place(relx=0.5, rely=0.10, anchor="center")

# Установка фона кнопки в черный цвет
btn_start.configure(bg='black')

video_label.place(x=0, y=0, relwidth=1, relheight=1)
command_label = tk.Label(root, text="Вы сказали: ", font=('Arial', 18), fg='white', bg='black')
command_label.pack(anchor='w', padx=10, pady=10)

btn_settings = tk.Button(text="⚙", font=('Arial', 12), command=open_settings, bg='black', fg='white',
                         highlightthickness=0, bd=0)
btn_settings.place(in_=video_label, relx=1, y=0, anchor='ne', width=30, height=30)

btn_close = tk.Button(text="Закрыть", font=('Arial', 12), command=close_program, bg='black', fg='white',
                      highlightthickness=0, bd=0)
btn_close.place(in_=video_label, relx=1, rely=1, anchor='se', width=60, height=30)

root.mainloop()
