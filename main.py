from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from random import randint
import openpyxl
import json


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument('user-data-dir=C:\\Google\\Chrome\\User Data') # УКАЖИТЕ ПУТЬ ГДЕ ЛЕЖИТ ВАШ ФАЙЛ. Советую создать отдельную папку.

workbook = openpyxl.load_workbook("number_list.xlsx")
worksheet = workbook.active


def load_numbers(): #Считывание всех номеров из xlsx
    list_number = [col[i].value for i in range(1, worksheet.max_row) for col in worksheet.iter_cols(1, worksheet.max_column) if col[i].value]
    return list_number

def click_new_chat():# Нажатие на кнопку новый чат
    browser.find_element(By.XPATH, '//*[@aria-label="Новый чат"]').click()
    sleep(randint(6, 11))

def entering_number(number):# Вставляем номер телефона
    browser.find_element(By.XPATH, '//*[@title="Текстовое поле поиска"]').send_keys(number)
    sleep(randint(6, 11))

def click_user_account():# Выбираем аккаунт "кому отправить"
    browser.find_element(By.XPATH, '//*[@tabindex="0" and @class="_199zF _3j691" and @role="button"]').click()
    sleep(randint(6, 11))

def entering_message(message): # Пишем текст
    browser.find_element(By.XPATH, '//*[@title="Введите сообщение"]').send_keys(message)
    sleep(randint(6, 11))

def click_send_message(): # Нажатие на кнопку отправить
    browser.find_element(By.XPATH, '//*[@aria-label="Отправить"]').click()
    sleep(randint(6, 11))

def click_back(): # Вернуться назад если пользователь не найден
    browser.find_element(By.XPATH, '//*[@aria-label="Назад"]').click()
    sleep(2)

result_json = {}

if __name__ == '__main__':
    list_number = load_numbers()
    message = 'Ваше сообщение'

    with webdriver.Chrome(options=options) as browser:
        browser.get('https://web.whatsapp.com/')
        sleep(20)
        for number in list_number:
            click_new_chat()
            entering_number(number)

            try:
                click_user_account()
                entering_message(message)
                click_send_message()
                print(f'Пользователь {number} Получил сообщение\n')
                result_json[number] = True

            except:
                click_back()
                print(f'Пользователь {number} Не найден\n')
                result_json[number] = '-------------------FALSE-------------------------'

            finally:
                with open('res.json', 'a', encoding='utf-8') as file:
                    json.dump(result_json, file, indent=4, ensure_ascii=False)
                    result_json = {}
