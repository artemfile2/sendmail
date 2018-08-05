import os
import re
import sys
import codecs
import smtplib
import fnmatch
import datetime
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.utils import formataddr
from urllib.parse import quote
from time import sleep
from future.backports.datetime import date


def send_mail(smtp_addr,
              fromaddr,
              password,
              toaddr,
              type_mail,
              sender=None,
              subject=None,
              text=None,
              data=None,
              ):
    from_addr = formataddr((sender, fromaddr))
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = toaddr
    msg['Subject'] = subject or ''

    if text:
        msg.attach(MIMEText(*text))

    if data:
        for filename, body in data.items():
            attachment = MIMEBase(mimetypes.guess_type(filename)[0].split('/')[0],
                                  mimetypes.guess_type(filename)[0].split('/')[1])
            attachment.set_payload(body)
            # attachment.set_payload(open(filename, "rb").read())
            encoders.encode_base64(attachment)

            if type_mail == 'rambler':
                filename = quote(filename)
                attachment.add_header('Content-Disposition',
                                      'attachment',
                                      filename=quote(filename))
            else:
                attachment.add_header('Content-Disposition',
                                      'attachment',
                                      filename=('utf_8', '', filename))


            msg.attach(attachment)

        server = smtplib.SMTP_SSL(smtp_addr[0], smtp_addr[1])
        try:
            server.login(fromaddr, password)
        except smtplib.SMTPAuthenticationError as err:
            sys.stderr.write(f"Проблема с отправкой письма. Причина: {err}")
        text = msg.as_string()

    try:
        server.send_message(msg)
    except (smtplib.SMTPRecipientsRefused, smtplib.SMTPSenderRefused) as err:
        sys.stderr.write(f"Проблема с отправкой письма. Причина: {err}")
    finally:
        server.quit()


def write_result_send(text, cnt):
    dt = datetime.datetime.now()
    with open('report.txt', "a+", encoding='utf-8-sig') as f:
        if cnt == 1:
            f.write('\n--------------------')
            f.write('\n\n'+str(dt) + '\n')
        f.write('\n'+ text)


def get_smtp(mail_adr):
    if mail_adr.find('rambler') >= 1:
        smtp_mail = 'smtp.rambler.ru'
        port_mail = 465
    elif mail_adr.find('mail') >= 1:
        smtp_mail = 'smtp.mail.ru'
        port_mail = 465
    elif mail_adr.find('yandex') >= 1:
        smtp_mail = 'smtp.yandex.ru'
        port_mail = 465
    elif mail_adr.find('gmail') >= 1:
        smtp_mail = 'smtp.gmail.com'
        port_mail = 465
    else:
        smtp_mail = 'any'
        port_mail = 465
    return smtp_mail, port_mail


def prepare():

    RootFoldersFile = open('from.lst')
    mail_from = RootFoldersFile.readline()
    mail_from.join(mail_from.replace('\n', ''))
    RootFoldersFile.close()

    mail_adr = mail_from.split('\\')[0]
    mail_pass = mail_from.split('\\')[1]

    mails_to = []
    mailslist = codecs.open('mails.lst', "r", "utf-8-sig")
    for line in mailslist:
        mails_to.append(line.replace('\n', ''))
    mailslist.close()

    i = 0
    for mail in mails_to:
        i += 1
        f_to = mail.split('\\')[0]
        m_to = mail.split('\\')[1]

        m_to = re.sub("^\s+|\n|\r|\s+$", '', m_to)
        f_to = re.sub("^\s+|\n|\r|\s+$", '', f_to)

        if m_to.find('rambler') >= 1:
            type_mail = 'rambler'
        else:
            type_mail = 'any'

        smtp_mail, port_mail = get_smtp(mail_adr)

        smtp_addr = smtp_mail, port_mail
        fromaddr = mail_adr
        password = mail_pass
        toaddr = m_to
        sender = "Макс-М экономический отдел"
        text = 'Здравствуйте, высылаем Вам 146 форму'

        data = {}
        path = r"C:\sendmail"
        fil = '*' + f_to + '*.xls'

        matches = []
        for root, dirnames, filenames in os.walk(path):
            for filename in fnmatch.filter(filenames, fil):
                matches.append(os.path.join(root, filename))
        with open(matches[0], "rb") as f:
            data[filename] = f.read()

        subject = filename

        send_mail(smtp_addr,
                  fromaddr,
                  password,
                  toaddr,
                  type_mail,
                  sender=sender,
                  subject=subject,
                  text=(text, 'plain', 'utf-8'),
                  data=data,
                  )

        now_date = datetime.datetime.now()
        now = now_date.strftime("%d.%m.%Y %H:%M:%S")

        text_msg = f"{i}| файл {f_to} на почтовый ящик {m_to} отправленн успешно... в {now}"
        write_result_send(text_msg, i)
        print(text_msg)
        sleep(60)

    input("Для выхода из программы нажмите Enter")


if __name__ == "__main__":
    prepare()