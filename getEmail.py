"""Получение файлов присылаемых из МАРС АРБИКОН из почты.

Программа для извлечения файлов с почтовых сообщений получаемых из АРБИКОНа.
Принцип работы: Обрабатывает из определенной папки куда складируются письма в почте (mail_folder_check)
каждое письмо. Проверяет есть ли в нем прикрепленные файлы с расширением .iso. Если есть то выгружает их
в папку на диске path_to_save, затем письмо перемещается в папку (mail_folder_unload) на почте с обработанными файлами
После выгрузки файлов из почты объединяет их в один iso файл. Затем формируется скрипт ImportMARC.IBF и запускается
АРМ "Администратор" с INI фалом в котором прописан путь до ImportMARC.IBF

Работает с почтой Яндекса, если переделывать на другую проверяейте синтаксис mail_folder_move
"""
# -*- coding: utf-8 -*-
import imaplib
import email
import os, time, sys
from email.header import decode_header
from operator import is_
from os.path import exists
from struct import pack
import subprocess
import socket

# Настройки
username = 'insert_your_mail@yandex.ru' # почта
password = 'password' # пароль от почты
path_to_save = 'D:\\ImportMARC\\' # путь куда сохраняются результаты выгрузки, объедененный iso файл, логи и IBF файл
mail_folder_check = 'arbicon' # папка в почте куда приходят письма
mail_folder_unload = 'Archives|2017' # папка куда будут перемещатся уже обработанные письма (Archives\2017)
irbis_path = 'D:\\IRBIS64\\' # путь где находится ИРБИС
irbis_param = 'IRBISA_IMPORT.INI' # ini файл для автозагрузки в ИРБИС в котором будет прописан BATCHFILE=D:\ImportMARC\ImportMARC.IBF

# Создаём файл для записи лога сообщений(Название файла: текущее время)
time_str = time.strftime("%d") + '-' + time.strftime("%m") + '-' + time.strftime("%Y") + '_' + time.strftime("%H") + '-' + time.strftime("%M") + '-' + time.strftime("%S")
path_to_save_iso_files = path_to_save + time_str + '\\'

# Создаем лог файл для записи сообщений
f_message = open(path_to_save+ '\\LOGS\\' + time_str+'.log', 'a')

# Функция для извлечения аттчей(только файлов с расширением iso) из писем
# возвращает количество обработанных писем и количество извлеченных файлов
# Соединяемся с сервером gmail.com через imap4_ssl
ymail = imaplib.IMAP4_SSL('imap.yandex.ru', '993')
ymail.login(username, password)

# Выбирает из папки входящие непрочитанные сообщения
typ, count = ymail.select(mail_folder_check)
f_message.write(count[0].decode() + " mail in folder " + mail_folder_check + "\n")

# Выводит количество непрочитанных сообщений в папке входящие
typ, unseen = ymail.status(mail_folder_check, '(UNSEEN)')
print("UNREAD :", unseen[0].decode(), " TOTAL: ", count[0].decode())

# Проверяем директорию куда будем писать
if not exists(path_to_save_iso_files):
    os.mkdir(path_to_save_iso_files)

# Главный блок
iso_files = []
try:
    typ, data = ymail.search(None, '(ALL)')	
    for i in data[0].split():
        # для проверки только на первых 10 письмах
        #if int(i) > 200:
        #    break
        typ, message = ymail.fetch(i, '(RFC822)')

        mail = email.message_from_bytes(message[0][1])
        parts = []
        ct_fields = []
        filenames = []
    
        for submsg in mail.walk():
            parts.append(submsg.get_content_type())
            ct_fields.append(submsg.get('Content-Type', ''))
            filenames.append(submsg.get_filename())
    
            if submsg.get_filename():
                if submsg.get_filename().endswith('.iso') :
                    with open(path_to_save_iso_files+submsg.get_filename(), 'wb') as new_file:
                        new_file.write(submsg.get_payload(decode=True))
                        iso_files.append(submsg.get_filename())
                        result, data = ymail.copy(i, mail_folder_unload)
                        if result == 'OK':
                            print('Письмо ' + i.decode() + ' удалено!')
                            ymail.store(i,'+FLAGS', '\\Deleted')
                        else:
                            print('Письмо '+ i.decode() +' не перемещено!')
                            f_message.write('Письмо '+ i.decode() + ' не перемещено!\n')
                print("Файл : ", submsg.get_filename())
                print("Длина файла : ", len(submsg.get_payload()))
                f_message.write('Unload from message ' + i.decode() + ' file - ' + submsg.get_filename() + '\n')
                time.sleep(1)
                # Permanently remove deleted items from selected mailbox
                ymail.expunge()
                time.sleep(0.5)
except Exception:
    print(sys.exc_info())
    f_message.write('Execept \n')
    pass
finally:
   # Permanently remove deleted items from selected mailbox
   ymail.expunge()

# Проверяем на всякий количество сохраненых файлов и считанных писем
if int(count[0]) != len(iso_files) :
    f_message.write('Unload ' + str(len(iso_files))  + ' from ' + count[0].decode() + "\n")
    print('Несоответствие количество выгруженых файлов количеству обработаных писем.')

# Отключаемся от сервера
ymail.close()
ymail.logout()

# получаем список файлов из диретории и пишем их все в один файл
iso_files_download = os.listdir(path_to_save_iso_files)
if len(iso_files_download) > 0:
    f_iso_file = open(path_to_save + time_str+'.iso', 'ab')
    f_message.write('Join ' + str(len(iso_files_download)) + ' iso files:' + '\n'.join(iso_files_download))
    for iso_file in iso_files_download:
        with open(path_to_save_iso_files+ '\\'+iso_file, 'rb') as f_read_iso:
            print('Open file...', iso_file)
            f_iso_file.write(f_read_iso.read())
            f_read_iso.close()
    f_iso_file.close()

   # Формируем скрипт .IBF с именем ISO файла выгруженных данных
    ImportIBFText= 'OpenDB MARC\nImportDB 0,marc_irb,0,1,' + path_to_save + time_str + '.iso\nCloseDB\nExit ' + path_to_save + time_str + '.txt'
    f_ibf_file = open(path_to_save + 'ImportMARC.IBF', 'w')
    f_ibf_file.write(ImportIBFText)
    f_ibf_file.close()
   # тут запускаем батник ирбиса для загрузки в КАТАЛОГ
    p = subprocess.Popen(irbis_path + 'irbisa.exe ' + irbis_param, shell=True, cwd=irbis_path)
    p.wait()
else:
    f_message.write('Nothing iso files ')

# Закрываем лог
f_message.close()
