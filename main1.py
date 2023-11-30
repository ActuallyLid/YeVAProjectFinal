import subprocess
import os
import sys

from PyQt5 import uic, QtCore, QtMultimedia
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog
from PyQt5.QtGui import QPixmap
import openpyxl
import openai
import threading

import win32com.client.dynamic
import pyttsx3
import speech_recognition as sr

import time

from pywinauto.application import Application

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

filename = 'РАСПИСАНИЕ.xlsx'
current_directory = os.getcwd()
file_path = os.path.join(current_directory, filename)

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

filename = 'РАСПИСАНИЕ.xlsx'
current_directory = os.getcwd()
file_path = os.path.join(current_directory, filename)

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
    'одиннадцатый': 11
}


def open_excel_file(file_path):
    try:
        subprocess.Popen(['start', 'excel', '/e', '/x', file_path], shell=True)
        return True
    except Exception as e:
        print(f"Ошибка при открытии файла Excel: {e}")
        return False


def find_cell_address(file_path, search_words):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for row in sheet.iter_rows():
        for cell in row:
            if cell.value in search_words:
                return cell

    return None


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


def voice(text):
    engine.say(text)
    engine.runAndWait()


def read_news():
    try:
        with open('news.txt', 'r', encoding='utf-8') as file:
            news_text = file.read()
            return news_text
    except Exception as e:
        print(f"Ошибка при чтении новостей: {e}")
        return None


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
        return word.lower()


class YeVAMainMenu(QMainWindow):
    def __init__(self):
        super(YeVAMainMenu, self).__init__()
        uic.loadUi('YeVAMainMenu.ui', self)
        self.voiceButton.clicked.connect(lambda: self.open_input())
        self.settingsBtn.clicked.connect(lambda: self.open_password())
        self.exitButton.clicked.connect(lambda: self.close_menu())

    def open_password(self):
        self.ui = PasswordDialogue()
        self.ui.show()

    def open_input(self):
        self.ui = InputMenu()
        self.ui.show()

    def close_menu(self):
        self.close()


class PasswordDialogue(QDialog):
    def __init__(self):
        super(PasswordDialogue, self).__init__()
        uic.loadUi('PasswordDialog.ui', self)
        self.pushButton.clicked.connect(lambda: self.check_password_settings())

    def check_password_settings(self):
        if self.lineEdit.text() == "pw87":
            try:
                self.ui = SettingsMenu()
                self.close()
                self.ui.show()
            except Exception as e:
                print(e)


class SettingsMenu(QMainWindow):
    def __init__(self):
        super(SettingsMenu, self).__init__()
        uic.loadUi('YeVaSettingsMenu.ui', self)
        self.quack('quack_5.mp3')

        pixmap1 = QPixmap('photo_2023-11-30_21-41-40.jpg')
        pixmap2 = QPixmap('photo_2023-11-30_21-51-50.jpg')

        self.cat_label.setPixmap(pixmap1)
        self.cat_label1.setPixmap(pixmap2)

        self.fileLoadButtonNews.clicked.connect(lambda: self.load_file_news())
        self.pushButtonNews.clicked.connect(lambda: self.save_changes_news())
        self.fileLoadButtonTable.clicked.connect(lambda: self.player.play())

    def load_file_news(self):
        with open('news.txt', 'r', encoding='utf8') as f:
            self.textBrowser.setText(f.read())

    def quack(self, filename):
        media = QtCore.QUrl.fromLocalFile(filename)
        content = QtMultimedia.QMediaContent(media)
        self.player = QtMultimedia.QMediaPlayer()
        self.player.setMedia(content)

    def save_changes_news(self):
        with open('news.txt', 'w', encoding='utf8') as f:
            f.write(self.textBrowser.toPlainText())
            self.textBrowser.setText('')


class InputMenu(QMainWindow):
    def __init__(self):
        super(InputMenu, self).__init__()
        uic.loadUi('YeVAInputMenu.ui', self)
        self.pushButton.clicked.connect(lambda: self.process_prog())

    def process_prog(self):
        try:
            def process():
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
                            break

                    if "открыть расписание" in command.lower() or "Открыть расписание" in command.lower():
                        print("Скажите номер и букву класса:")
                        self.textEdit_Listen.setText("Скажите номер и букву класса:")
                        voice("Пожалуйста, скажите номер и букву класса.")
                        r = sr.Recognizer()
                        with sr.Microphone() as source:
                            audio = r.listen(source)
                        try:
                            search_words = r.recognize_google(audio, language="ru-RU").replace(" ", "").split(',')
                            search_words = [convert_to_russian(word) for word in search_words]

                            print("Вы сказали:", search_words)

                            self.textEdit_Listen.setText(last_command)

                        except sr.UnknownValueError:
                            response = "Извините, не удалось распознать ключевые слова"
                            print("Извините, не удалось распознать ключевые слова.")

                            continue
                        search_words = [f"Класс - {convert_to_number_or_word(word)}" for word in search_words]

                        cell = find_cell_address(file_path, search_words)
                        if cell is not None:
                            maximize_excel_window()
                            navigate_to_cell(cell)

                            response = (f"Расписание ( {convert_to_number_or_word(search_words[0])})"
                                        f" было успешно развернуто и перемещено к ячейке")
                        else:
                            response = (f"  ({convert_to_number_or_word(search_words[0])})"
                                        f" не были найдены в расписании. Пожалуйста, повторите запрос.")

                    elif ("новости" in command.lower() or "Новости школы" in command.lower() or "новости школы"
                          in command.lower() or "Новости" in command.lower() or "Новости в школе" in command.lower()
                          or "новости в школе" in command.lower()):
                        response = read_news()
                        self.textEdit_Listen.setText(last_command)
                    else:
                        openai.api_key = "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"
                        response = "извините, повторите вопрос "

                    response = response.replace("Класс - ", "")

                    print("ИИ:", response)

                    voice(response)
                    self.textEdit_Listen.setText(last_command)

                    self.textEdit_AI.setText(response)

                    break

            threading.Thread(target=process()).start()
        except Exception as e:
            print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = YeVAMainMenu()
    ex.show()
    sys.exit(app.exec())
