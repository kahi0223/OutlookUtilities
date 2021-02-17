import win32com.client as win32
from datetime import datetime
import os
from config import MailConfig
from config import PathConfig

class Emailer:

    def __init__(self) -> None:
        self.outlook = win32.Dispatch('outlook.application') # lanching Outlook
        self.inbox = self.outlook.GetNamespace("MAPI").GetDefaultFolder(6) # Inbox
        self.mailer = self.outlook.CreateItem(0)
    
    def create_new_mail(self,
                        subject: str,
                        body: str,
                        to: list,
                        cc: list,
                        attach_paths: list,
                        display: bool = True) -> None:
        """ create new email and save it. not send by automatically
        :param subject: subject of email
        :param body: body text of email
        :param to: list of TO
        :param cc: list of CC
        :param attach_paths: list of files to attach
        :param display: pop up or not after created and saved email
        """
        print("----Creating New Mail:", subject)
        self.mailer.To = ";".join(to)
        self.mailer.CC = ";".join(cc)
        self.mailer.Subject = subject
        self.mailer.HtmlBody = body

        for path in attach_paths:
            self.mailer.Attachments.Add(path)
            print(">> Attached:", path)
        if display:
            self.mailer.Display()
        self.mailer.Save()

    def search_mail_and_save_attachments(self,
                         subject: str,
                         save_path:str,
                         folder_name: str = None) -> list:
        """save attachment file of email automatically.
        find email from sepecified subject.
        if found several same subects, newest only
        :param subject: target subject to find
        :param save_path: folder path to save attachment
        :param folder_name: target sub folder to find email. default(None) means Inbox
        :return: list of saved attachment paths if found. [] if not found.
        """
        attachment_paths = []
        print("----Finding Mail:", subject)
        if folder_name:
            folder = self.inbox.Folders(folder_name)
        else:
            folder = self.inbox

        # seaching email
        for mail in folder.Items:
            filter = mail.Subject.startswith(subject)

            if filter:
                body = mail.Body
                body = body.replace("\r\n", "<br>")
                
                attachments = mail.Attachments
                
                # saving attachments
                if attachments:
                    for attachment in attachments:
                        attachment.SaveAsFile(save_path)
                        attachment_paths.append(save_path)
                    print("> found:", attachment_paths)
        if not attachment_paths:
            print("> None")
        return attachment_paths