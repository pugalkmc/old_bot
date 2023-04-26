import os
import re
import time
from datetime import datetime, timedelta
from typing import List

import openpyxl
import firebase_admin
from firebase_admin import credentials, db
from openpyxl.styles import Alignment
from telegram import *
from telegram.ext import *
from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import get_column_letter
import os
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from openpyxl import Workbook

import google.auth

# Set up the YouTube API client
api_key = 'AIzaSyDj1k7YAx4HGCFtp_k7x_eB5-u2wUSi2oI'
youtube = build('youtube', 'v3', developerKey=api_key)

li = [""]

print(len(li))

cred = credentials.Certificate("kit-pro-f4b0d-firebase-adminsdk-mhzrf-8a07acab1c.json")
firebase_admin.initialize_app(cred, {
    "databaseURL": "https://kit-pro-f4b0d-default-rtdb.firebaseio.com/"
})

# Set up the Telegram bot

bot = Bot(token="6208523031:AAFfOb97T6Wml0pZUagE56A_MZDpCpUXZJk")


def start(update, context):
    message = update.message
    chat_id = message.chat_id
    bot.sendMessage(chat_id=chat_id,
                    text="Hi! I'm your Telegram bot. I'll collect messages and links from PoolSea Group")


def collect_message(update, context):
    message = update.message
    username = message.chat.username
    chat_id = message.chat_id
    chat_type = message.chat.type
    text = message.text

    if chat_type == "private":
        if "get " in message.text and len(text) > 6:
            text = text.replace("get ", "")
            selva_sheet(update=update, context=context, date=text)
        elif username not in ["srikanth084", "Jellys04", "Cryptomaker143", "Shankar332", "Royce73", "Balaharishb",
                              "SaranKMC", "pugalkmc", "SebastienKulec"]:
            bot.sendMessage(chat_id=chat_id, text="You have no permission to use this bot")
            return
        if "spreadsheet admin" in text:
            int_org = text.replace("spreadsheet", "")
            save_to_spreadsheet(update, context, admin=int_org)

        elif "spreadsheet" in message.text and len(message.text) > 12:
            text = text.replace("spreadsheet ", "")
            save_to_spreadsheet(update=update, context=context, date=text)

        # scraping video for query
        elif "video " in text:
            new = []
            video_duration = 'medium'
            text = text.replace("video ", "")
            # search_terms = text.split(" ")
            min_subscribers = 100
            max_results = 10000

            # Search for videos with the given hashtag and at least min_subscribers subscribers
            videos = []
            next_page_token = None
            while len(videos) < max_results:
                search_response = youtube.search().list(
                    q=text,
                    type="video",
                    part="id,snippet",
                    order="viewCount",
                    videoDuration=video_duration,
                    maxResults=max_results,
                    pageToken=next_page_token
                ).execute()

                for search_result in search_response.get("items", []):
                    video_id = search_result["id"]["videoId"]
                    video_response = youtube.videos().list(
                        id=video_id,
                        part="id,snippet,statistics",
                    ).execute()
                    videos.append(f"https://www.youtube.com/watch?v={video_id}")

                if "nextPageToken" in search_response:
                    next_page_token = search_response["nextPageToken"]
                else:
                    break

            print(f"Retrieved {len(videos)} videos.")

            # Create a spreadsheet with the collected channels
            wb = Workbook()
            ws = wb.active
            new = videos
            for i in new:
                ws.append([i])
            wb.save(f"{text}.xlsx")
            bot.sendDocument(chat_id=chat_id, document=open(f"{text}.xlsx", "rb"))


        # getting channel link for the videos in li list
        elif "start " in text:
            # Use the API to get the channel ID and URL of the video
            name = text.replace("start ", "")
            urls = []
            num = 1
            for i in li:
                try:
                    if (num % 100 == 0):
                        print(f"Completed: {num}")
                    num += 1
                    id = i.replace("https://www.youtube.com/watch?v=", "")
                    response = youtube.videos().list(
                        part='snippet',
                        id=id
                    ).execute()
                    if 'items' in response and len(response['items']) > 0:
                        # response is not empty
                        channel_id = response['items'][0]['snippet']['channelId']
                        channel_response = youtube.channels().list(
                            part='statistics',
                            id=channel_id
                        ).execute()
                        if 'items' in channel_response and len(channel_response['items']) > 0:
                            # channel response is not empty
                            subscriber_count = int(channel_response['items'][0]['statistics']['subscriberCount'])
                            if subscriber_count >= 100:
                                channel_url = f'https://www.youtube.com/channel/{channel_id}'
                                urls.append(channel_url)


                except HttpError as error:
                    print(f'An error occurred: {error}')
                    channel_url = ''

            # Create a spreadsheet with the collected channels
            wb = Workbook()
            ws = wb.active
            urls = list(set(urls))
            for i in urls:
                ws.append([i])
            wb.save(f"{name}.xlsx")
            bot.sendDocument(chat_id=chat_id, document=open(f"{name}.xlsx", "rb"))


    elif chat_type == "group" or chat_type == "supergroup":
        # Only process messages from specific users in personal chat
        collection_name = datetime.now().strftime("%Y-%m-%d")
        message_id = message.message_id
        message_date_ist = (datetime.now() + timedelta(hours=5, minutes=30)).strftime(
            "%H:%M:%S")  # Convert datetime to IST timezone
        text = message.text
        if chat_id == -827109122:
            db.reference(f'selva/{collection_name}/{message_id}').set({
                'username': username,
                'text': text,
                'time': message_date_ist,
                'message_id': message_id
            })

        if chat_id not in [-1001588000922, -1588000922] or username not in ["srikanth084", "Jellys04", "Cryptomaker143",
                                                                            "Shankar332", "Royce73",
                                                                            "Balaharishb",
                                                                            "LEO_sweet_67",
                                                                            "SaranKMC"]:
            return

        # Store message data in Firebase Realtime Database
        db.reference(f'messages/{collection_name}/{message_id}').set({
            'username': username,
            'text': text,
            'time': message_date_ist,
            'message_id': message_id
        })


admins_list = [1155684571, 814546021, 1291659507]


def selva_sheet(update, context, admin=None, date=None):
    collection_name = date if date else datetime.now().strftime("%Y-%m-%d")
    bot.sendMessage(chat_id=1292480260, text=f"request received for date: {collection_name}")
    # collection_name = (datetime.now() + timedelta(hours=5, minutes=30)).strftime("%Y-%m-%d")

    # Get all the messages from the database for a specific date
    messages = db.reference(f'selva/{collection_name}').get() or {}

    if (len(messages) == 0 or messages is None):
        bot.sendMessage(chat_id=1292480260, text="No message found for today")
    # Create a new Excel workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    # Write the headers
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 18

    ws['A1'] = 'Username'
    ws['B1'] = 'Message Text'
    ws['C1'] = 'IST Time'

    new_li = []
    msg_dict = {}
    row = 2
    for message_id, message_data in messages.items():
        username = message_data.get('username')
        text = message_data.get('text')
        time = message_data.get('time')
        new_li.append([username, text, time])
        ws.cell(row=row, column=1).value = username
        ws.cell(row=row, column=2).value = text
        ws.cell(row=row, column=3).value = time
        row += 1

        if username in msg_dict:
            msg_dict[username] += 1
        else:
            msg_dict[username] = 1
    wb.save(f"{collection_name}.xlsx")
    row = 1
    for i in msg_dict:
        ws.cell(row=row, column=5).value = i
        ws.cell(row=row, column=6).value = msg_dict[i]
        row += 1

    wb.save(f"{collection_name}.xlsx")

    bot.sendDocument(chat_id=1292480260, document=open(f"{collection_name}.xlsx", "rb"))


def save_to_spreadsheet(update, context, admin=None, date=None):
    collection_name = date if date else datetime.now().strftime("%Y-%m-%d")
    # collection_name = (datetime.now() + timedelta(hours=5, minutes=30)).strftime("%Y-%m-%d")

    # Get all the messages from the database for a specific date
    messages = db.reference(f'messages/{collection_name}').get() or {}
    if (len(messages) == 0 or messages is None):
        bot.sendMessage(chat_id=1291659507, text="Hey! ,No message found for today")

    # Create a new Excel workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # Write the headers
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 20
    ws['A1'] = 'Username'
    ws['B1'] = 'Message Link'
    ws['C1'] = 'Message Text'
    ws['D1'] = 'IST Time'
    ws['F1'] = 'Username'
    ws['G1'] = 'Message Count'

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
            row += 1
    msg = ""
    for i in username_counts:
        msg += f"{i} - {username_counts[i]['count']}\n"
    bot.sendMessage(chat_id=update.message.chat_id, text=f"Total Messages: {len(messages.items())}\n\n"
                                                         f"{msg}")
    if admin is not None and "group" in admin:
        bot.sendMessage(chat_id=-1001586628789, text=f"Count for {collection_name}: {len(messages.items())}\n\n"
                                                     f"{msg}")

    ws["F1"] = "Usernames"
    ws["G1"] = "Count"

    ws['F2'] = 'Jellys04'
    ws['F3'] = 'Cryptomaker143'
    ws['F4'] = 'Shankar332'
    ws['F5'] = "Royce73"
    ws['F6'] = "Balaharishb"
    ws['F7'] = "SaranKMC"
    ws['F8'] = "srikanth084"

    # set the formula in cell G2
    for row in range(2, 9):
        username = ws.cell(row=row, column=6).value  # Get the username from Column F
        count = '=COUNTIF(A:A,"*' + username + '*")'  # Construct the formula
        ws.cell(row=row, column=7).value = count

    # Save the Excel workbook
    file_name = f"{collection_name}.xlsx"
    wb.save(file_name)
    if admin is None:
        bot.sendDocument(chat_id=update.message.chat_id, document=open(file_name, 'rb'))
    if admin is not None and "admin" in admin:
        if "group" in admin:
            bot.sendDocument(chat_id=-1001586628789, document=open(file_name, "rb"))
        for i in admins_list:
            bot.sendDocument(chat_id=i, document=open(file_name, "rb"))


def main():
    updater = Updater(token="6208523031:AAFfOb97T6Wml0pZUagE56A_MZDpCpUXZJk", use_context=True)
    dp = updater.dispatcher
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CommandHandler("spreadsheet", save_to_spreadsheet))
    dp.add_handler(MessageHandler(Filters.text, collect_message))
    updater.start_polling()


main()
