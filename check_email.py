import imaplib
import ssl
import email
import telegram
import logging.handlers
import logging
import json, time, os, datetime

#SSL Context Setting
SSL_CONTEXT = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
SSL_CONTEXT.set_ciphers('DEFAULT@SECLEVEL=1')

settings = {}
telegramed_mails = []
telegram_bot = None
logger = logging.getLogger()

def ConnectMailSvr():
    try:
        mail = imaplib.IMAP4_SSL(
            settings['mail_server'],
            port=settings['mail_port'],
            ssl_context=SSL_CONTEXT
        )
        mail.login(settings['mail_userid'], settings['mail_passwd'])
        mail.select("INBOX", readonly=True)
        _, num_ids_data = mail.search(None, '(UNSEEN)')
        mail_count = 0
        global telegramed_mails
        for message_id in num_ids_data[0].split():
            _, msg_data = mail.fetch(message_id, '(RFC822)')
            email_info = GetContents(msg_data)
            if email_info not in telegramed_mails:
                send_telegram(email_info)
                telegramed_mails.append(email_info)
                mail_count += 1

        if mail_count == 0:
            logger.info('-')
        mail.close()
    except Exception as e:
        logger.warn('[예외발생]', e)

def GetContents(data):
    try:
        raw_email = data[0][1]
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)
        mail_from = email.utils.parseaddr(email_message['From'])
        from_text = get_decoded_text(mail_from[0])
        from_addr = mail_from[1]
        subject = get_decoded_text(email_message['Subject'])
        mail_date_p = datetime.datetime.strptime(email_message['Date'],"%a, %d %b %Y %H:%M:%S %z")
        mail_date = mail_date_p.strftime('%Y.%m.%d %H:%M:%S')

        email_info = '발신자명: {}<{}>\n발신일시: {}\n메일제목: {}'.format(
            from_text, from_addr, mail_date, subject)
        return email_info
    except Exception as e:
        logger.warn('[예외발생]', e)

def get_decoded_text(encoded_text):
    if len(encoded_text) > 0:
        decoded_text, encoding = email.header.decode_header(encoded_text)[0]
        decoded_text = decoded_text.decode(encoding)
    else:
        decoded_text = ''
    return decoded_text

def send_telegram(email_info):
    logger.info(email_info)
    telegram_bot.sendMessage(
        chat_id = settings['telegram_chatid'],
        text='[메일이 도착했습니다!]\n{}'.format(email_info)
    )

def set_os_timezone():
    os.environ['TZ'] = settings['logging_timezone']
    time.tzset()

def initialize_logger():
    global logger
    set_os_timezone()
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        '[%(asctime)s | %(levelname)s] %(message)s',
        '%Y.%m.%d %H:%M:%S'
    )
    timedfilehandler = logging.handlers.TimedRotatingFileHandler(
        filename='check_email.log',
        when='midnight',
        interval=1,
        encoding='utf-8'
    )
    timedfilehandler.setFormatter(formatter)
    timedfilehandler.suffix = "%Y%m%d"
    logger.addHandler(timedfilehandler)

def get_settings():
    global settings, telegram_bot
    with open('check_email.json') as json_file:
        settings = json.load(json_file)
    telegram_bot = telegram.Bot(token = settings['telegram_token'])

def check_telegramed_mails():
    global telegramed_mails
    MAX_TELEGRAMED_MAILS = settings['max_telegramed_mails']
    if len(telegramed_mails) > MAX_TELEGRAMED_MAILS:
        send_telegram('{}개 이상의 메일이 아직 확인되지 않습니다.'.format(
            MAX_TELEGRAMED_MAILS))
        telegramed_mails = []

def CheckMailLoop():
    while True:
        check_telegramed_mails()
        ConnectMailSvr()
        time.sleep(settings['sleep_interval'])

if __name__ == "__main__":
    try:
        get_settings()
        initialize_logger()
        CheckMailLoop()
    except Exception as e:
        logger.warn('[예외발생]', e)
