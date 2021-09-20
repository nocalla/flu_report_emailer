import sys

import win32com.client as win32


class EmailHandler:
    def __init__(self, account_str) -> None:
        self.outlook = self.get_outlook()
        self.account = self.select_account(account_str)

    def get_outlook(self):
        try:
            outlook = win32.Dispatch("outlook.application")
            return outlook
        except:
            # Exit the program if Outlook is not open
            print("Error: Outlook is not open - emails cannot be created.")
            sys.exit(1)

    def select_account(self, search: str) -> object:
        """select the account to send the email from

        Args:
            search (str): desired search term for account name

        Returns:
            object: outlook account object for the desired account
        """
        outlook = self.outlook
        accounts = outlook.Session.Accounts
        for account in accounts:
            if search in str(account):
                from_account = account
                break
        print("Account selected: {}".format(from_account))
        return from_account

    def create_email(self, account: object, mail_details: list) -> object:
        """create an Outlook email object based on mail_details list and save as draft under specified Outlook account

        Args:
            account (object): specified Outlook account
            mail_details (list): email details in form [to_address, subject, body_text]

        Returns:
            object: outlook mail object
        """
        outlook = self.outlook
        mail = outlook.CreateItem(0)
        mail.To = mail_details[0]
        mail.Subject = mail_details[1]
        mail.HTMLBody = mail_details[2]
        # set the "send from" account using arcane methods
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
        # save a draft
        mail.Save()
        # open the email in a popup window (disabled in favour of saving a draft)
        # mail.Display(False)

