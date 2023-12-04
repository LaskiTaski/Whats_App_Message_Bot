from selenium import webdriver
from selenium.webdriver.common.by import By
from random import randint
import openpyxl
import json
import time

wookbook = openpyxl.load_workbook("number_list.xlsx")
worksheet = wookbook.active

input_number = '//*[@id="app"]/div/div[2]/div[2]/div[1]/span/div/span/div/div[1]/div[2]/div[2]/div/div[1]/p'
input_message = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p'

btn_new_chat = '//*[@id="app"]/div/div[2]/div[3]/header/div[2]/div/span/div[3]/div'
btn_user_account = '//*[@id="app"]/div/div[2]/div[2]/div[1]/span/div/span/div/div[2]/div[2]'
btn_send_message = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button'
btn_back = '//*[@id="app"]/div/div[2]/div[2]/div[1]/span/div/span/div/div[1]/div[2]/button'

def load_numbers(): #Считывание всех номеров из xlsx
    list_number = [col[i].value for i in range(1, worksheet.max_row) for col in worksheet.iter_cols(1, worksheet.max_column) if col[i].value]
    return list_number

def click_new_chat():# Нажатие на кнопку новый чат
    browser.find_element(By.XPATH, btn_new_chat).click()
    time.sleep(randint(6, 11))

def entering_number(number):# Вставляем номер телефона
    browser.find_element(By.XPATH, input_number).send_keys(number)
    time.sleep(randint(6, 11))

def click_user_account():# Выбираем аккаунт "кому отправить"
    browser.find_element(By.XPATH, btn_user_account).click()
    time.sleep(randint(6, 11))

def entering_message(message): # Пишем текст
    browser.find_element(By.XPATH, input_message).send_keys(message)
    time.sleep(randint(6, 11))

def click_send_message(): # Нажатие на кнопку отправить
    browser.find_element(By.XPATH, btn_send_message).click()
    time.sleep(randint(6, 11))

def click_back(): # Вернуться назад если пользователь не найден
    browser.find_element(By.XPATH, btn_back).click()
    time.sleep(2)

result_json = {}
if __name__ == '__main__':
    list_number = load_numbers()
    message = 'Ваше сообщение'

    with webdriver.Chrome() as browser:
        browser.get('https://web.whatsapp.com/')
        time.sleep(20)
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

