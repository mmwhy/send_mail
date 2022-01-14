import os
from st2common.runners.base_action import Action
import win32com.client as win32
__all__ = [
    'SendExchangeMailAction'
]
receiver='2920405578@qq.com'
subject='Mail Test'
HTMLBody='''
        <H2>Hello, This is a test mail.</H2>
        Hello Guys. 
        '''
class SendExchangeMailAction(Action):

    def run(self,receiver,subject,HTMLBody):
        outlook = win32.Dispatch('Outlook.Application')

        mail_item = outlook.CreateItem(0)  # 0: olMailItem

        mail_item.Recipients.Add(receiver)
        mail_item.Subject = subject

        mail_item.BodyFormat = 2  # 2: Html format
        mail_item.HTMLBody = HTMLBody
        mail_item.Send()
