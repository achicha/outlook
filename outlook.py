from win32com.client import Dispatch
from telegram import Bot
import configparser
import logging
from datetime import datetime, timedelta


# Enable logging
logging.basicConfig(filename='outlook.txt', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    level=logging.INFO)
logger = logging.getLogger(__name__)
logger.warning('restart')  # log startup


class Outlook:
    def __init__(self, pers_folder, update_time):
        # "6" refers to the index of a folder
        self.inbox = Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6)
        self.personal_folder = self.inbox.Folders(pers_folder)
        self.messages = self.personal_folder.Items
        self.time = update_time

    def last_messages(self, number):
        emails = []
        if isinstance(self.time, str):
            self.time = float(self.time)
        for i in range(len(self.messages) - number, len(self.messages)):
            if self.messages[i].CreationTime.timestamp() > self.time:
                emails.append((self.messages[i].SenderName, self.messages[i].Subject))
                self.time = self.messages[i].CreationTime.timestamp()
        return emails

    def updated(self):
        return str(self.time)


class Telegram:
    def __init__(self):
        self.bot = Bot(token=bot_access_token)

    def send_message(self, msg_text):
        self.bot.sendMessage(chat_id=chat_id, text=msg_text)


class Controller:
    pass


if __name__ == '__main__':
    # load config from file
    config = configparser.ConfigParser()
    config.read('./config.ini')
    bot_access_token = config['Telegram']['access_token']
    chat_id = config['Telegram']['chat_id']
    outlook_folder = config['Outlook']['private']
    # if last update is emtpy then we going to use the default value
    last_update = config['Outlook']['last_update'] or (datetime.now() - timedelta(1)).timestamp()

    # outlook
    outlook = Outlook(outlook_folder, last_update)
    mails = outlook.last_messages(3)
    text = ('\n'.join([str(m)[1:-2] for m in mails]))
    print(text)

    # update config
    config['Outlook']['last_update'] = outlook.updated()

    with open('./config.ini', 'w') as configfile:
        config.write(configfile)

    # telegram
    tele = Telegram()
    tele.send_message(text)

