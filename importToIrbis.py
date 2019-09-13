"""Скрипт для загрузки .iso файла в БД МАРС

После выгрузки файлов из почты объединяет их в один iso файл. Затем формируется скрипт ImportMARC.IBF и запускается
АРМ "Администратор" с INI фалом в котором прописан путь до ImportMARC.IBF

Работает с почтой Яндекса, если переделывать на другую проверяейте синтаксис mail_folder_move
"""
# -*- coding: utf-8 -*-
import os, time
from operator import is_
from os.path import exists
from struct import pack
import subprocess

# Настройки
path_to_save = 'D:\\ImportMARC\\' # путь куда сохраняются результаты выгрузки, объедененный iso файл, логи и IBF файл
irbis_path = 'D:\\IRBIS64\\'
irbis_param = 'IRBISA_IMPORT.INI'

# Создаём файл для записи лога сообщений(Название файла: текущее время)
time_str = time.strftime("%d") + '-' + time.strftime("%m") + '-' + time.strftime("%Y") + '_' + time.strftime("%H") + '-' + time.strftime("%M") + '-' + time.strftime("%S")
# путь к каталогу где собраны все iso файлы для объединения
path_to_save_iso_files = path_to_save + '06-09-2018_11-38-10' + '\\' 

# Создаем лог файл для записи сообщений
f_message = open(path_to_save + time_str+'.log', 'a')

# Проверяем директорию куда будем писать
if not exists(path_to_save_iso_files):
    f_message.write('Not found path :' + path_to_save_iso_files)
    exit(1)

# получаем список файлов из диретории и пишем их все в один файл
iso_files_download = os.listdir(path_to_save_iso_files)
if len(iso_files_download) > 0:
    f_iso_file = open(path_to_save + time_str+'.iso', 'ab')
    f_message.write('Folder ' + path_to_save_iso_files + '\n')
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
