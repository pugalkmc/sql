import logging
import os
import re
import time
from datetime import datetime, timedelta

import openpyxl
import pytz
from openpyxl.styles import Alignment
from telegram import *
from telegram.ext import *
import datetime as date_mod
import mysql.connector

token = "5123712096:AAF9i762sqwDi76OvHPqD8YWNHjMGIWFiko"
# Set up the MySQL connection

bot = Bot(token=token)

# Replace the placeholders with your own credentials
host = 'pugalkmc.mysql.pythonanywhere-services.com'
database = 'pugalkmc$poolsea'
user = 'pugalkmc'
password = 'pugalsaran143'

# Connect to the database
try:
    conn = mysql.connector.connect(host=host, database=database, user=user, password=password)
    cursor = conn.cursor()
    print('Connected to MySQL database on PythonAnywhere')
    insert_query = f"INSERT INTO messages (username, message_id, message_date, message_text) VALUES ('john_doe', 12345, '2022-02-23 12:34:56', 'Hello, world!')"
    cursor.execute(insert_query)
    conn.commit()
    if cursor.rowcount > 0:
        print("Insert successful")
    else:
        print("Insert failed")
    bot.sendMessage(chat_id=1291659507, text="inserted")

except mysql.connector.Error as e:
    print(f'Error connecting to MySQL database: {e}')

# Set up the Telegram bot


def start(update, context):
    update.message.reply_text("Hi! I'm your Telegram bot. I'll collect messages and links from this group.")
    

def collect_message(update, context):
    message = update.message
    username = message.from_user.username
    chat_id = message.chat_id
    chat_type = message.chat.type

    if chat_type == "private":
        if username not in ["Jellys04", "Cryptomaker143", "Shankar332" , "Royce73", "Balaharishb", "LEO_sweet_67", "SaranKMC", "pugalkmc"]:
            bot.sendMessage(chat_id=chat_id, text="You have no permission to use this bot")
            return

        if "/spreadsheet" in message.text and len(message.text) > 12:
            save_to_spreadsheet(date_mod.datetime.now().strftime("%Y-%m-%d"))

    elif chat_type == "group":
        # Only process messages from specific group and users
        if chat_id not in [-1001869015888,-1001605512981, -1001588000922] or username not in ["Jellys04", "Cryptomaker143", "Shankar332" , "Royce73", "Balaharishb", "LEO_sweet_67", "SaranKMC", "pugalkmc"]:
            bot.sendMessage(chat_id=1291659507, text="returned")
            return

        # Only process messages from specific users in personal chat
        collection_name = date_mod.datetime.now().strftime("%Y-%m-%d")
        message_id = message.message_id
        message_date = message.date
        message_text = message.text
        insert_query = f"INSERT INTO messages (username, message_id, message_date, message_text) VALUES ('{username}', '{message_id}', '{message_date}', '{message_text}')"
        cursor.execute(insert_query)
        db.commit()
        if cursor.rowcount > 0:
            print("Insert successful")
        else:
            print("Insert failed")



def save_to_spreadsheet(update=None,context=None, date=None):
    collection_name = date_mod.datetime.now().strftime("%Y-%m-%d") if date is None else date

    # Get all the messages from the database
    select_query = f"SELECT * FROM messages WHERE message_date >= '{collection_name}' AND message_date < '{collection_name} 23:59:59'"
    cursor.execute(select_query)
    messages = [{'username': row[0], 'message_id': row[1], 'message_date': row[2], 'message_text': row[3]} for row in cursor.fetchall()]

    # Create a new Excel workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active

        # Write the headers
    ws.column_dimensions['A'].width = 13
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 100
    ws['A1'] = 'Username'
    ws['B1'] = 'Message ID'
    ws['C1'] = 'Message Date'
    ws['D1'] = 'Message Text'

    # Write the data to the worksheet
    for i, message in enumerate(messages, start=2):
        ws.cell(row=i, column=1, value=message['username'])
        ws.cell(row=i, column=2, value=message['message_id'])
        ws.cell(row=i, column=3, value=message['message_date'].strftime('%Y-%m-%d %H:%M:%S'))
        ws.cell(row=i, column=4, value=message['message_text'])

    # Save the workbook
    wb.save('chat_history.xlsx')
    bot.sendDocument(chat_id=1291659507, document=open('chat_history.xlsx', "rb"))
    
def run_bot():
    updater = Updater(token=token, use_context=True)
    dp = updater.dispatcher
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CommandHandler("spreadsheet", save_to_spreadsheet))
    dp.add_handler(MessageHandler(Filters.text, collect_message))
    updater.start_polling()


if __name__ == "__main__":
    run_bot()
