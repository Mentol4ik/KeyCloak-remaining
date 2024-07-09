import smtplib
import urllib
import os
import sys
import pandas as pd
import string
import openpyxl as xl
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from platform import python_version
import openpyxl as xl



class Task:
    def __init__(self, user_id, uuid, name, email, login, new_pass):
        self.user_id = user_id
        self.uuid = uuid
        self.name = name
        self.email = email
        self.login = login
        self.new_pass = new_pass
        self.status = 0  #1: success, 2:error_keycloak, 3:error_notify

    def __str__(self):
        return f'user_id =,{self.user_id},uuid=,{self.uuid},login =, {self.login}, mail =, {self.email}, name=,{self.name}, pass =, {self.new_pass}, status=,{self.status}'


# Parse file and return results
def parse_xls_file(file_name):
    results = []
    _list = pd.read_excel('Passwords_test.xlsx', sheet_name='Лист1')
    row_count = _list.shape[1]
    for i in range(1, row_count):
        user_id = i
        uuid = _list["ID"][i]
        name = _list["Name"][i]
        mail = _list["Email"][i]
        new_pass = _list["New_password"][i]
        login = _list["Login"][i]
        results.append(
            Task(
                user_id=user_id,
                uuid = uuid,
                login = login,
                email = mail,
                name = name,
                new_pass = new_pass
            )
            )
    return results


'''
def  sending():
    server = 'smtp.yandex.ru'#Для отправки НЕ с Яндекса изменить на smtp....
    user = 'mentolhik@yandex.ru'#логин 
    password = 'spkqvfjwhuvxmjdt'#пароль от аккаунта
    recipients = 'mentolhik@yandex.ru'#Здесь указываются адреса всех кому должно придти письмо
    sender = 'mentolhik@yandex.ru'#Здесь указывается адрес отправителя
    subject = 'Обновление пароля на сайте'#Здесь указывается тема письма
    text = 'Здравствуйте, '#+ appeal + ', Ваш пароль на сайте обновился. Ваш новый пароль: '+ newpassword #Здесь записывается текст самого сообщения
    html = '<html><head></head><body><p>' + text + '</p></body></html>'
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = 'Sityacionniy centr<' + sender + '>'
    msg['To'] = ', '.join(recipients)
    msg['Reply-To'] = sender
    msg['Return-Path'] = sender

    part_text = MIMEText(text, 'plain')
    part_html = MIMEText(html, 'html')
    msg.attach(part_text)
    msg.attach(part_html)
    mail = smtplib.SMTP_SSL(server)
    mail.login(user, password)
    mail.sendmail(sender, recipients, msg.as_string())
    mail.quit()
    #spkqvfjwhuvxmjdt
sending()
'''


def resset_password(task) -> bool:
    print(f'Resset password for user with id=,{task.uuid}')

    keycloak_admin = KeycloakAdmin(server_url="https://auth.erdc.ru/auth/",
                                   username='admin',
                                   password='123',
                                   realm_name="test-sc-password-skip",
                                   # user_realm_name="only_if_other_realm_than_master",
                                   # client_secret_key="test-sc-password-skip",
                                   verify=True)
    response = keycloak_admin.set_user_password(user_id=task.uuid, password=task.new_pass,
                                                temporary=True)
    return True


def send_notify_to_user(task) -> bool:
    print(f'Sending notify for user with id=,{task.uuid}')
    return True



def main():

    file_name = os.path.basename(r'/Users/matvejsvarev/PycharmProjects/KOD/Passwords_test.xlsx')
    tasks = parse_xls_file(file_name)
    shrink_tasks = tasks[1:].copy()
    for task in shrink_tasks:
        parse_xls_file(file_name)
        print(task.results, 'Running ends successfully')
        print(task)


if __name__ == '__main__':
    sys.exit(main())
