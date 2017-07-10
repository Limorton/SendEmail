# -*- coding: utf-8 -*-
"""
Created on Mon Jul 10 18:38:39 2017
这是一个包含自动发送邮件类的模块，它定义了一个实现发送邮件功能的类，和一个使用该类实现发送带附件的邮件的函数
@author: Limorton
""" 
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
#from email.utils import COMMASPACE
from email.mime.text import MIMEText
#from email.header import Header
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
import configparser
import smtplib
import os

class SendEmail:
    def __init__(self,email_config_path,email_attachment_path):
        """
        从配置文件获取发送邮件的配置信息：
            SMTP邮箱服务器地址和端口：host, port
            用户名和密码(授权码)：name,pwd
            发件人和收件人：sender,receiver
        附件的文件路径：email_attachment_path
        """
        config = configparser.ConfigParser()
        config.read(email_config_path)
        self.attachment_path = email_attachment_path
        # 第三方 SMTP 服务配置
        self.host = config.get('SMTP','host')                # 设置服务器
        self.port = config.get('SMTP','port')                # 设置端口
        self.username = config.get('SMTP','login_username')  # 用户名(邮箱)
        self.password = config.get('SMTP','login_password')  # 密码(或授权码，比如QQ邮箱的授权码)
        # 邮件的发件人和接收人的配置
        self.sender = config.get('TEXT','sender')
        self.receiver = config.get('TEXT','receiver')
        
    def login(self):
        """
        login in to smtp e-mail host
        """
        self.smtp = smtplib.SMTP_SSL(self.host,self.port)
        try:
            self.smtp.login(self.username,self.password)
            print('登陆邮箱成功！')
        except:
            print('登陆邮箱失败！')
            #raise AttributeError('Can not login smtp!')
    
    def send(self,email_title,email_content):
        """
        send e-mail
        """
        msg = MIMEMultipart()
        msg['From'] = self.sender
        msg['To'] = self.receiver
        msg['Subject'] = email_title
        receiver = self.receiver.split(';')
        content = MIMEText(email_content,_charset = 'gbk')
        msg.attach(content)
        # 附件处理
        for attachment_name in os.listdir(self.attachment_path):
            attachment_file = os.path.join(self.attachment_path,attachment_name)
            
            with open(attachment_file,'rb') as attachment:
                if 'application' == 'text':
                    attachment = MIMEText(attachment.read(), _subtype = 'octet-stream', _charset='GB2312')
                elif 'application' == 'image':
                    attachment = MIMEImage(attachment.read(), _subtype = 'octet-stream')
                elif 'application' == 'audio':
                    attachment = MIMEAudio(attachment.read(), _subtype = 'octet-stream')
                else:
                    attachment = MIMEApplication(attachment.read(), _subtype = 'octet-stream')
            attachment.add_header('Content-Disposition','attachment',filename = ('gbk', '', attachment_name))
            msg.attach(attachment)
        self.smtp.sendmail(self.sender,receiver,msg.as_string())
    
    def quit(self):
        self.smtp.quit()

# 使用上述类，定义函数实现发送邮件
def send_Email():
    import time
    time_form = '%Y-%m-%d_%A'
    #current_time = str(time.strftime(time_form))
    
    email_config_path = './config/emailConfig.ini'               # 此处定义配置文件路径
    email_attachment_path = './attachments'                      # 此处定义附件所在文件夹
    
    email_title = 'wuzhuti2'                                     # 邮件的主题
    email_content = '这封邮件是上一封邮件的追加，包含附件在内'     # 邮件的内容
    
    email = SendEmail(email_config_path,email_attachment_path)
    email.login()
    email.send(email_title,email_content)
    email.quit()
    print('邮件发送成功! 时间: ',str(time.strftime(time_form)))
     
if __name__ == '__main__':
    send_Email()