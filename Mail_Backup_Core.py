import datetime
import email
import imaplib
import os
import quopri
import re
import base64
import time
from email.header import decode_header, make_header
import configparser
from tqdm import tqdm

def backup_config():
    # 导入config.ini文件信息
    config_data = configparser.ConfigParser()
    config_data.read("config.ini", encoding='utf-8')

    User = config_data.get('Setting', 'User')
    Passwd = config_data.get('Setting', 'Passwd')

    Email_box = config_data.get('Setting', 'Email_box')
    Batch_size = int(config_data.get('Setting', 'Batch_size'))

    Imap_url = config_data.get('Email', 'Imap_url')
    Port = int(config_data.get('Email', 'Port'))

    return User, Passwd, Email_box, Batch_size, Imap_url, Port


def mail_login(Imap_url, Port, User, Passwd):
    global BIT_mail
    try:
        BIT_mail = imaplib.IMAP4(host=Imap_url, port=Port) if Port == 143 else imaplib.IMAP4_SSL(host=Imap_url,
                                                                                                 port=Port)
        BIT_mail.login(User, Passwd)
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ' Successful login')
    except Exception as e:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ' Login failed')
        print("ErrorType : {}, Error : {}".format(type(e).__name__, e))

    return BIT_mail


def encoded_words_to_text(encoded_words):
    encoded_word_regex = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}='
    try:
        text = str(make_header(decode_header(encoded_words)))
    except:
        try:
            charset, encoding, encoded_text = re.match(encoded_word_regex, encoded_words).groups()
            if encoding == 'B':
                byte_string = base64.b64decode(encoded_text)
                text = byte_string.decode(charset)
            elif encoding == 'Q':
                byte_string = quopri.decodestring(encoded_text)
                text = byte_string.decode(charset)
        except:
            text = '无主题'
    return text


def save_mailbox(Imap_url, Port, User, Passwd, email_box, Batch_size):
    BIT_mail = mail_login(Imap_url, Port, User, Passwd)
    status, data = BIT_mail.select(email_box, readonly=True)
    if status == 'OK':
        print("Processing mailbox: ", email_box)
        status_email_uid, email_uid = BIT_mail.search(None, "ALL")
        if status_email_uid != 'OK':
            print("No messages found!")
            return
        else:
            print(f"{len(email_uid[0].split())} emails in the folder")

            # 抓取邮件
            if Batch_size > int(data[0].decode()):
                Batch_size = int(data[0].decode())

            with tqdm(total=len(email_uid[0].split()), desc='Saving ' + email_box, mininterval=0.3) as pbar:
                for epoch in range(1, (len(email_uid[0].split())-1) // Batch_size + 2):
                    if epoch * Batch_size < len(email_uid[0].split()):
                        num = ('%s:%s') % (epoch * Batch_size - (Batch_size - 1), epoch * Batch_size)
                    else:
                        num = ('%s:*') % (epoch * Batch_size - (Batch_size - 1))

                    status_email_data, email_data = BIT_mail.fetch(num, '(RFC822)')
                    if status_email_data != 'OK':
                        print("ERROR getting message " + num)
                        return
                    # print("Writing message " + num)
                    for id in range(0, len(email_data), 2):
                        my_msg = email.message_from_bytes(email_data[id][1])
                        Subject = encoded_words_to_text(my_msg['Subject'])
                        From = str(make_header(decode_header(my_msg['From']))).replace("<", "- ").replace(">", "")
                        Date = time.strftime("%Y-%m-%d",
                                             time.localtime(
                                                 email.utils.mktime_tz(email.utils.parsedate_tz(my_msg['Date']))))

                        savedir = './' + email_box + '/'
                        if not os.path.isdir(savedir):
                            os.makedirs(savedir)

                        filename_raw = (Date + "_" + From + "_" + Subject).replace("@", "#").replace(".", "_").replace(" ", "").replace("-", "_").replace("\"", "_")
                        subst = ''
                        regex = r"[:.。,，（）=+、\\\/!！@%$￥*?\"'|“”\[\]；<>]"
                        filename = re.sub(regex, subst, filename_raw, 0, re.MULTILINE).replace('\r\n', '').replace('\n', '').replace('__', '_')

                        savedir_filename = '%s%s.eml' % (savedir, filename)
                        # str((id // 2 + 1)).zfill(len(str(len(email_uid[0].split())))) + " - " +  # 邮件编号
                        with open(savedir_filename, 'wb') as f:
                            # print("Writing message " + str((id // 2 + 1) + (epoch - 1) * Batch_size) + "/" + str(
                            #     len(email_uid[0].split())))
                            f.write(email_data[id][1])

                        pbar.update(1)
    else:
        print("ERROR: Unable to open mailbox ", status)

    BIT_mail.close()
    print('Mailbox closed.')
    BIT_mail.logout()
    print('Logout')
