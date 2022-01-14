import os
from st2common.runners.base_action import Action
import win32com.client as win32
__all__ = [
    'SendExchangeMailAction'
]

class SendExchangeMailAction(Action):

    def run(self,receiver,subject,HTMLBody):
        outlook = win32.Dispatch('Outlook.Application')

        mail_item = outlook.CreateItem(0)  # 0: olMailItem

        mail_item.Recipients.Add(receiver)
        mail_item.Subject = subject

        mail_item.BodyFormat = 2  # 2: Html format
        mail_item.HTMLBody = HTMLBody
        mail_item.Send()
