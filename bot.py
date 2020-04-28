import telebot
import datetime
import os
import pickle
from telebot import apihelper
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request

bot = telebot.TeleBot('telegrambotapi')


# использую прокси для общения с ботом для отладки на локальной машине.
# на хостинге это необязательно
apihelper.proxy = {'https':'proxy'}

spreadsheet_id = 'spreadsheet_id'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


# добавляем несколько кнопок
keyboard1 = telebot.types.ReplyKeyboardMarkup()
keyboard1.row('in', 'out', 'time left', 'lunch')

def main():
    global values_input, service
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

def get_time_left():
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range = range_time_left).execute()
    return result.get('values', [])

def write_in():
    day = datetime.date.today().day
    time_in = str((datetime.datetime.utcnow()+datetime.timedelta(hours=3)).time()).split('.')[0]
    range_in = 'C' + str(day+1)

    response_date = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='RAW',
        range=range_in,
        body=dict(
            majorDimension='ROWS',
            values=[[time_in]])
    ).execute()
    print('Entry time added successfully')

def write_out():
    day = datetime.date.today().day
    time_out = str((datetime.datetime.utcnow()+datetime.timedelta(hours=3)).time()).split('.')[0]
    range_out = 'D' + str(day+1)

    response_date = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='RAW',
        range=range_out,
        body=dict(
            majorDimension='ROWS',
            values=[[time_out]])
    ).execute()
    print('Exit time added successfully')

def exclude_lunch():
    day = datetime.date.today().day
    range_lunch_bool = 'F' + str(day+1)

    response_date = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='RAW',
        range=range_lunch_bool,
        body=dict(
            majorDimension='ROWS',
            values=[[1]])
    ).execute()
    print('Lunch excluded successfully')

@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, 'Привет, ты написал мне /start', reply_markup=keyboard1)

@bot.message_handler(content_types=['text'])
def send_text(message):
    msg = message.text
    msg_low = message.text.lower()
    if msg_low == 'date':
        bot.send_message(message.chat.id, str(day))
    elif msg_low == 'in':
        write_in()
        bot.send_message(message.chat.id, 'Добро пожаловать на работу!')
    elif msg_low == 'out':
        write_out()
        bot.send_message(message.chat.id, "Поздравляю, ты вышел с работы!")
    elif msg_low == 'lunch':
        exclude_lunch()
        bot.send_message(message.chat.id, "Надеюсь, было вкусно...")
    elif msg_low == 'time left':
        bot.send_message(message.chat.id, *get_time_left())
    elif msg_low == 'привет' or msg_low == 'здравствуй' or msg_low == 'hi' or msg_low == 'hello':
        bot.send_message(message.chat.id, 'Привет!')
    elif msg_low == 'пока' or msg_low == 'bye' or msg_low == 'goodbye' or msg_low == 'до свидания':
        bot.send_message(message.chat.id, 'До свидания!')
    else:
        bot.send_message(message.chat.id, 'Я не умею отвечать на '+ str(msg)+'"')

@bot.message_handler(content_types=['sticker'])
def sticker_id(message):
    bot.send_message(message.chat.id, 'Классный стикер!')

main()

range_time_left = 'G2'

bot.polling()
