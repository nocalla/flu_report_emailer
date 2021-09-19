import win32com.client as win32
import csv
import sys


def get_outlook():
    try:
        outlook = win32.Dispatch("outlook.application")
        return outlook
    except:
        print("Error: Outlook is not open.")
        sys.exit(1)


def get_data(file):  # TODO
    # obtain data from our file
    data = []  # placeholder DEBUG
    with open(file) as f:
        csv_reader = csv.reader(f, delimiter=",")
        for row in csv_reader:
            process_row(row)

    return data  # placeholder DEBUG


def process_row(row):
    """ 
    ['Date Dispensed', 'Patient Name', 'Street', 'Town or City', 'Birth Date',
    'PPSN No', 'Gender', 'Qty', 'Script Dispensed As', 'Directions Expanded',
    'Contract GP Name', 'Contract GP Address']
    """
    pass


def create_email(account, address, subject, body):
    # creates an email
    outlook = get_outlook()
    mail = outlook.CreateItem(0)
    mail.To = address
    mail.Subject = subject
    mail.Body = body
    # set the "send from" account using arcane methods
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
    # mail.Display(False)  # maybe save draft instead of this. handy for debug though
    mail.Save()
    return mail


def format_body(name, patients):
    # composes the message text
    # TODO: html version instead to allow for formatting?
    greeting = "Dear {},\n".format(name)
    general_body = "For your information, the below patients of yours were recently vaccinated in our pharmacy. Vaccine details, including batch and expiry, are below.\n"
    patient_details = "{}".format(patients)
    sign_off = "Kind regards,\n"
    body_string = greeting + general_body + patient_details + sign_off

    return body_string


def select_account(search):
    # select the account to send the email from
    outlook = get_outlook()
    accounts = outlook.Session.Accounts
    for account in accounts:
        if search in str(account):
            from_account = account
            break
    print("Account selected: {}".format(from_account))
    return from_account


def create_recipient_list():  # TODO
    # create our list of recipients based on the data we have
    recipient_list = {
        "Adrian Test": ["Patient A", "Patient B"],
        "Beatrice Test": ["Patient C", "Patient D"],
    }  # debug values
    return recipient_list


def create_emails(account, recipient_list):
    # for each recipient in list, create & save a draft email
    for recipient in recipient_list:
        print("Composing email for {}".format(recipient))
        body = format_body(recipient, recipient_list[recipient])
        subject = "Vaccine Report - {}".format(recipient)
        address = "test@address"
        mail = create_email(account, address, subject, body)


def main():
    # find the file with data
    data = get_data("Flu Vacc Report.csv")
    # parse the data, creating list?

    # select correct account
    account = select_account("healthmail")
    # generate list of recipients
    recipient_list = create_recipient_list()
    # create emails
    create_emails(account, recipient_list)


if __name__ == "__main__":
    main()
