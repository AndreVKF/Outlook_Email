from win32com.client import Dispatch

import win32com.client
from Email import Email

######### Variaveis Locais #########
Email = Email(destinatario='thiago@octante.com.br', assunto='Teste')

######### Connect to Outlook #########

outlook = Dispatch('Outlook.Application').GetNamespace('MAPI')
inbox = outlook.GetDefaultFolder('6')

######### Create Email #########

const = win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch('Outlook.Application')
newMail = obj.CreateItem(olMailItem)

######### Email Content #########

newMail.Subject = Email.mail_assunto
newMail.HtmlBody = Email.mail_body

######### Set Recipients to Mail #########
newMail.To = Email.mail_destinatario
newMail.CC = 'middle@octante.com.br'

######### Show Email before send #########
newMail.display()
