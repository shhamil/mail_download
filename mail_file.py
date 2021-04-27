import datetime
import email
import os
import getpass
import re
import time
from itertools import groupby
from imaplib import IMAP4_SSL
import base64
MAIL_HOST = "httmlimap.mail.ru"
MAIL_PORT = 993
user_txt = 'last_user.txt'
check_last_user = os.path.exists('last_user.txt')
MAIL_USER = ''
MAIL_PASSWORD = ''
if check_last_user:
    yn = input('Войти используя последнюю авторизованную почту(y/n):')
    if yn == 'y' or yn == 'Y':
        with open(user_txt, "r", newline="") as file:
            user_info = file.read().splitlines()
            MAIL_USER = user_info[0]
            MAIL_PASSWORD = user_info[1]
    elif yn == 'n' or yn == 'N':
        MAIL_USER = input('Enter mail: ')
        MAIL_PASSWORD = getpass.getpass('Enter password: ')
        with open(user_txt, "w", newline="") as file:

            file.write(MAIL_USER + '\n')
            file.write(MAIL_PASSWORD)
            file.close()
else:
    MAIL_USER = input('Enter mail: ')
    MAIL_PASSWORD = getpass.getpass('Enter password: ')
    with open(user_txt, "w", newline="") as file:

        file.write(MAIL_USER + '\n')
        file.write(MAIL_PASSWORD)
        file.close()

DATE_FOR_DOWNLOAD = input('Введите дату до которой необходимо скачать вложения(01 Jan 2000): ')
DATETIME_FORMAT_STRING = r'\d\d \w\w\w \d\d\d\d \d\d:\d\d:\d\d'
DATE_FORMAT_STRING = r'\d\d \w\w\w \d\d\d\d'


def get_date_from_message_for_filename(msg):
    res = []
    for elem in msg:
        match = re.search(DATETIME_FORMAT_STRING, elem)
        res.append(match[0] if match else 'Not found')

    res = [el for el, _ in groupby(res)]
    result = ''
    for r in res:
        if r != 'Not found':
            result += r
            break

    return result


def get_date_from_message_for_path(msg):
    res = []
    for elem in msg:
        match = re.search(DATE_FORMAT_STRING, elem)
        res.append(match[0] if match else 'Not found')

    res = [el for el, _ in groupby(res)]
    result = ''
    for r in res:
        if r != 'Not found':
            result += r
            break

    return result


def write_filename_in_csv(filename, count):
    txt_name = "FileName.txt"
    try:
        file = open(txt_name, 'a')
        file.write(filename + '\n')
        file.close()
    finally:
        print('Скачан:', count)


def check_file_was_download(fname):
    txt_name = "FileName.txt"
    try:
        with open(txt_name, "r", newline="") as file:
            df = file.read().splitlines()
            if fname in df:
                file.close()
                return True
    except FileNotFoundError:
        file = open(txt_name, 'w')
        file.close()
        return False


def download_files():
    connection = IMAP4_SSL(host=MAIL_HOST, port=MAIL_PORT)
    connection.login(user=MAIL_USER, password=MAIL_PASSWORD)
    status, msgs = connection.select('INBOX')
    assert status == 'OK'
    i = 1
    typ, data = connection.search(None, 'ALL')
    for msg_id in reversed(data[0].split()):
        status, msg_data = connection.fetch(msg_id, '(RFC822)')
        mail = email.message_from_bytes(msg_data[0][1])
        if mail.is_multipart():
            filelist = []
            for part in mail.walk():
                if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                    continue
                content_type = part.get_content_type()
                filename = part.get_filename()
                transfer_encoding = part.get_all('Content-Transfer-Encoding')
                if transfer_encoding and transfer_encoding[0] == 'base64':
                    if filename is not None:
                        filename_parts = filename.split('?')
                        if len(filename_parts) != 1:
                            filename = base64.b64decode(filename_parts[3]).decode(filename_parts[1])
                if filename:
                    filename = get_date_from_message_for_filename(mail.values()).replace(' ', '-').replace(
                        ':', '-') + filename
                    path = './xls/' + get_date_from_message_for_path(mail.values()).replace(' ', '-') + '/' + (
                        mail.values()[1].replace('<', '')).replace('>', '') + '/'
                    if datetime.datetime.strptime(get_date_from_message_for_path(mail.values()), '%d %b %Y') \
                            <= datetime.datetime.strptime(DATE_FOR_DOWNLOAD, '%d %b %Y'):
                        return 'Все файлы до ' + DATE_FOR_DOWNLOAD + ' скачаны'
                    if not os.path.exists(path):
                        os.makedirs(path)
                    if 'pdf' in content_type:
                        filename = filename + '.pdf'
                    is_download = check_file_was_download(filename)
                    if is_download:
                        return 'Все файлы скачаны: ' + datetime.datetime.today().strftime('%H:%M - %d %b %Y')
                    filelist.append(filename)
                    write_filename_in_csv(filename, i)
                    i += 1
                    with open(path + filename, 'wb') as new_file:
                        new_file.write(part.get_payload(decode=True))
                    print()
    connection.close()
    connection.logout()


while True:
    print(download_files())
    time.sleep(300)
