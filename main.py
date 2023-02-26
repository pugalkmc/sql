import os
import re
import time
from datetime import datetime, timedelta
from typing import List

import openpyxl
import pytz
import firebase_admin
from firebase_admin import credentials, db
from openpyxl.styles import Alignment
from telegram import *
from telegram.ext import *

cred = credentials.Certificate("kit-pro-f4b0d-firebase-adminsdk-mhzrf-8a07acab1c.json")
firebase_admin.initialize_app(cred, {
    "databaseURL": "https://kit-pro-f4b0d-default-rtdb.firebaseio.com/"
})

# Set up the Telegram bot

bot = Bot(token="6208523031:AAFfOb97T6Wml0pZUagE56A_MZDpCpUXZJk")


def start(update, context):
    message = update.message
    chat_id = message.chat_id
    bot.sendMessage(chat_id=chat_id, text="Hi! I'm your Telegram bot. I'll collect messages and links from PoolSea Group")


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
            save_to_spreadsheet(admin="yes", update=update, context=context)
        elif "spreadsheet" in message.text and len(message.text) > 12:
            save_to_spreadsheet(datetime.now().strftime("%Y-%m-%d"), update=update, context=context)

    elif chat_type == "group" or chat_type == "supergroup":
        if chat_id not in [-1001588000922] or username not in ["Jellys04", "Cryptomaker143", "Shankar332", "Royce73",
                                                               "Balaharishb",
                                                               "LEO_sweet_67",
                                                               "SaranKMC", "pugalkmc"]:
            return

        # Only process messages from specific users in personal chat
        collection_name = datetime.now().strftime("%Y-%m-%d")
        message_id = message.message_id
        message_date_ist = (datetime.now() + timedelta(hours=5, minutes=30)).strftime("%Y-%m-%d %H:%M:%S")  # Convert datetime to IST timezone
        message_text = message.text

        # Store message data in Firebase Realtime Database
        db.reference(f'messages/{collection_name}/{message_id}').set({
            'username': username,
            'text': message_text,
            'time': message_date_ist,
            'message_id': message_id
        })



def save_to_spreadsheet(admin="yes", update=None, context=None, date=None):
    collection_name = date if date else datetime.now().strftime("%Y-%m-%d")

    # Get all the messages from the database for a specific date
    messages = db.reference(f'messages/{collection_name}').get() or {}

    # Create a new Excel workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # Write the headers
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20
    ws['A1'] = 'Username'
    ws['B1'] = 'Message Link'
    ws['C1'] = 'Message Text'
    ws['D1'] = 'IST Time'
    ws['E1'] = 'Count'
    ws['F1'] = 'Unique Usernames'

    # Write the data
    row = 2
    username_counts = {}
    for message_id, message_data in messages.items():
        username = message_data.get('username')
        text = message_data.get('text')
        time = message_data.get('time')
        link = f'https://t.me/poolsea/{message_id}'
        
        if username:
            if username in username_counts:
                username_counts[username]['count'] += 1
            else:
                username_counts[username] = {'count': 1, 'total': 0}

            ws.cell(row=row, column=1).value = username
            ws.cell(row=row, column=2).value = link
            ws.cell(row=row, column=3).value = text
            ws.cell(row=row, column=4).value = time
            ws.cell(row=row, column=5).value = 1
            ws.cell(row=row, column=6).value = username
            row += 1
    
    # Write the unique usernames and their message counts to column E
    row = 2
    for username, counts in username_counts.items():
        ws.cell(row=row, column=5).value = counts['count']
        row += 1

    # Save the Excel workbook
    file_name = f'{collection_name}.xlsx'
    wb.save(file_name)
    bot.sendDocument(chat_id=1291659507, document=open(file_name, 'rb'))
    if admin == "yes":
        bot.sendDocument(chat_id=814546021, document=open(file_name, "rb"))


        
def main():
    updater = Updater(token="6208523031:AAFfOb97T6Wml0pZUagE56A_MZDpCpUXZJk", use_context=True)
    dp = updater.dispatcher
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CommandHandler("spreadsheet", save_to_spreadsheet))
    dp.add_handler(MessageHandler(Filters.text, collect_message))
    updater.start_polling()


main()
