import logging
import os
import re
import time
from datetime import datetime, timedelta

import openpyxl
import pytz
from openpyxl.styles import Alignment
from telegram import *
from telegram.bot import Bot
from telegram.ext import *
import datetime as date_mod
import mysql.connector

token = "6208523031:AAH0jWiZr8FOEZ_1xyarUg0-liaMUcDn3uw"
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

except mysql.connector.Error as e:
    print(f'Error connecting to MySQL database: {e}')


# Set up the Telegram bot


def start(update, context):
    update.message.reply_text("Hi! I'm your Telegram bot. I'll collect messages and links from PoolSea Group")


def collect_message(update, context):
    message = update.message
    username = message.from_user.username
    chat_id = message.chat_id
    chat_type = message.chat.type
    text = message.text

    if chat_type == "private":
        if username not in ["Jellys04", "Cryptomaker143", "Shankar332", "Royce73", "Balaharishb", "LEO_sweet_67",
                            "SaranKMC", "pugalkmc"]:
            bot.sendMessage(chat_id=chat_id, text="You have no permission to use this bot")
            return
        if "spreadsheet admin" == text:
            save_to_spreadsheet(admin="yes")
        elif "spreadsheet" in message.text and len(message.text) > 12:
            save_to_spreadsheet(date_mod.datetime.now().strftime("%Y-%m-%d"))

    elif chat_type == "group" or chat_type == "supergroup":
        if chat_id not in [-1001588000922] or username not in ["Jellys04", "Cryptomaker143", "Shankar332", "Royce73",
                                                               "Balaharishb",
                                                               "LEO_sweet_67",
                                                               "SaranKMC", "pugalkmc"]:
            return

        # Only process messages from specific users in personal chat
        collection_name = date_mod.datetime.now().strftime("%Y-%m-%d")
        message_id = message.message_id
        message_date = message.date.strftime("%Y-%m-%d %H:%M:%S")  # Convert datetime to string
        message_text = message.text
        insert_query = f"INSERT INTO messages (username, message_id, message_text, message_date) VALUES ('{username}', '{message_id}', '{message_text}', '{message_date}')"
        cursor.execute(insert_query)
        conn.commit()


def save_to_spreadsheet(admin="yes", update=None, context=None, date=None):
    collection_name = date if date else datetime.now().strftime("%Y-%m-%d")

    # Get all the messages from the database for a specific date
    select_query = f"SELECT username, message_id, message_text, message_date FROM messages WHERE DATE(message_date) = '{collection_name}'"
    cursor.execute(select_query)
    messages = [{'username': row[0], 'message_id': row[1], 'message_text': row[2], 'message_date': row[3]} for row in
                cursor.fetchall()]

    user_counts = {}
    for message in messages:
        if message['username'] in user_counts:
            user_counts[message['username']]['count'] += 1
        else:
            user_counts[message['username']] = {'count': 1, 'total': 0}

    # Calculate the total number of messages for the day
    total_messages = sum([user_counts[username]['count'] for username in user_counts])

    # Update the total message count for each user
    for username in user_counts:
        user_counts[username]['total'] = total_messages

    # Create a new Excel workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # Write the headers and user message counts
    ws.column_dimensions['A'].width = 13
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 100
    ws['A1'] = 'Username'
    ws['B1'] = 'Message Link'
    ws['C1'] = 'Message Text'
    ws['D1'] = 'Message Date'
    ws['E1'] = """=QUERY(ARRAYFORMULA(LOWER(B:B)),"SELECT Col1, COUNT(Col1) WHERE Col1 <> '' GROUP BY Col1 LABEL COUNT(Col1) 'Count'",1)"""
    for i, (username, counts) in enumerate(user_counts.items(), start=2):
        ws.cell(row=i, column=1, value=username)
       # ws.cell(row=i, column=5, value=counts['count'])
        # ws.cell(row=i, column=6, value=counts['total'])

    # Write the data to the worksheet
    for i, message in enumerate(messages, start=len(user_counts) + 2):
        ws.cell(row=i, column=1, value=message['username'])
        ws.cell(row=i, column=2, value=f"https://t.me/poolsea/{message['message_id']}")
        ws.cell(row=i, column=3, value=message['message_text'])
        ws.cell(row=i, column=4, value=message['message_date'])

    # Freeze the top row of the sheet
    ws.freeze_panes = ws['A2']

    # Save the workbook
    wb.save('chat_history.xlsx')

    bot.sendDocument(chat_id=1291659507, document=open('chat_history.xlsx', "rb"))
    if admin == "yes":
        bot.sendDocument(chat_id=1155684571, document=open('chat_history.xlsx', "rb"))


def main():
    updater = Updater(token=token, use_context=True)
    dp = updater.dispatcher
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CommandHandler("spreadsheet", save_to_spreadsheet))
    dp.add_handler(MessageHandler(Filters.text, collect_message))
    updater.start_polling()


main()


