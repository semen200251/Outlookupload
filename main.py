import os
import datetime
import win32com.client
import logging
import config

logging.basicConfig(
    filename='outlook.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S' 
)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

try:
    folder = outlook.GetDefaultFolder(6).Folders(config.name_folder_email)
    logging.info(f'Папка "{config.name_folder_email}" найдена')
except Exception as e:
    logging.error(f'Ошибка при получении папки "{config.name_folder_emai}": {e}')
    raise SystemExit()

messages = folder.Items

today = datetime.datetime.today()
start_date = today - datetime.timedelta(days=1)

messages = messages.Restrict("[ReceivedTime] >= '" + start_date.strftime('%m/%d/%Y %H:%M %p') + "'")
save_folder = config.path_to_folder_PC
if os.path.exists(save_folder):
    for message in messages:
        try:
            attachments = message.Attachments
            for attachment in attachments:
                if attachment.FileName.endswith('.xlsx'):
                    attachment.SaveAsFile(os.path.join(os.getcwd(), attachment.FileName))
                    logging.info(f'Файл "{attachment.FileName}" успешно скачан')
        except Exception as e:
            logging.error(f'Ошибка при скачивании вложений сообщения: {e}')
            continue
else:
    logging.error(f'Ошибка в пути к папке на компьютере {config.path_to_folder_PC}')
    raise SystemExit()
