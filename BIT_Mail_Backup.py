import datetime
import email
import imaplib
import os
import quopri
import re
import base64
import time
from email.header import decode_header, make_header


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


def save_mailbox(Imap_url, Port, User, Passwd, email_box, batch_size):
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
            for epoch in range(1, len(email_uid[0].split()) // batch_size + 2):
                if epoch * batch_size < len(email_uid[0].split()):
                    num = ('%s:%s') % (epoch * batch_size - (batch_size - 1), epoch * batch_size)
                else:
                    num = ('%s:*') % (epoch * batch_size - (batch_size - 1))

                status_email_data, email_data = BIT_mail.fetch(num, '(RFC822)')
                if status_email_data != 'OK':
                    print("ERROR getting message " + num)
                    return
                # print("Writing message " + num)
                for id in range(0, len(email_data), 2):
                    my_msg = email.message_from_bytes(email_data[id][1])
                    Subject_raw = encoded_words_to_text(my_msg['Subject'])
                    subst = ''
                    regex = r"[:.。,，（）=+、\\\/!！@%$￥*?\"'|“”\[\]；<>]"
                    Subject = re.sub(regex, subst, Subject_raw, 0, re.MULTILINE).replace('\r\n', '').replace('\n', '')
                    Date = time.strftime("%Y-%m-%d",
                                         time.localtime(
                                             email.utils.mktime_tz(email.utils.parsedate_tz(my_msg['Date']))))

                    savedir = './' + email_box + '/'
                    if not os.path.isdir(savedir):
                        os.makedirs(savedir)

                    filename = '%s%s.eml' % (savedir, Date + " - " + Subject)
                    # str((id // 2 + 1)).zfill(len(str(len(email_uid[0].split())))) + " - " +  # 邮件编号
                    with open(filename, 'wb') as f:
                        print("Writing message " + str((id // 2 + 1) + (epoch - 1) * batch_size) + "/" + str(
                            len(email_uid[0].split())))
                        f.write(email_data[id][1])
    else:
        print("ERROR: Unable to open mailbox ", status)

    BIT_mail.close()
    print('Mailbox closed.')
    BIT_mail.logout()
    print('Logout')


def main():
    # 必填项
    User = ""
    Passwd = ""

    # 选填项
    email_box = "Inbox"  # 收件箱
    # Sent 已发送
    # Draft 草稿箱
    # Trash 已删除
    batch_size = 20  # 受限于网络带宽，不要同时获取大量邮件，建议不超过50封一次

    Imap_url = 'mail.bit.edu.cn'
    Port = 143

    save_mailbox(Imap_url, Port, User, Passwd, email_box, batch_size)


if __name__ == "__main__":
    main()
