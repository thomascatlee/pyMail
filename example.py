#!/usr/bin/env python2.7
# -*- coding: utf-8 -*
import pyMail
import getpass
import time
import os
import filename

# 推荐使用qq邮箱
username = raw_input("User: ")
passwd = getpass.getpass()
smtp_server = raw_input("smtp server: ")
imap_server = raw_input("imap server: ")

# 初始化发送邮件类
sml = pyMail.SendMailDealer(username, passwd, smtp_server, 25)
# 设置邮件信息
sml.setMailInfo(username, u'标题', u'正文', 'plain', './README.md')
# 发送邮件
sml.sendMail()

# 初始化接收邮件类
rml = pyMail.ReceiveMailDealer(username, passwd, imap_server)

account_dir = os.getcwd() + '/email/' + username

rml.select('INBOX')

box_dir = account_dir + '/' + 'INBOX'

os.makedirs(box_dir, exist_ok=True)
os.chdir(box_dir)

for num in rml.search(None,"All")[1][0].split(b' '):
    if num != b'':
        mailInfo = rml.getMailInfo(num)

        print(num)

        os.chdir(box_dir)

        #保存邮件到eml文件
        name = filename.getWindowsName(mailInfo['subject'] + '-' + mailInfo['date'])
        with open((name + '.eml'), 'wb') as fp:
            fp.write(mailInfo['eml'])

        #如果有附件则建立文件夹
        if len(mailInfo['attachments']) > 0:
            email_dir = os.path.join(box_dir, name)
            os.makedirs(email_dir, exist_ok=True)
            os.chdir(email_dir)

        # 遍历附件列表
        for attachment in mailInfo['attachments']:
            fileob = open(attachment['name'], 'wb')
            fileob.write(attachment['data'])
            fileob.close()