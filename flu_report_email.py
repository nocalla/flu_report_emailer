import win32com.client as win32
import csv
import sys


class Recipient:
    def __init__(self, name, address, email, patient_entry):
        self.name = name
        self.address = address
        self.email = email
        self.patients = []
        self.patients.append(patient_entry)

    def add_patient_entry(self, patient_entry):
        self.patients.append(patient_entry)


class PatientEntry:
    def __init__(self, data_row):
        self.date = data_row[0]
        self.patient = data_row[1]
        self.address = "{}, {}".format(data_row[2], data_row[3])
        self.dob = data_row[4]
        self.ppsn = data_row[5]
        self.gender = data_row[6]
        self.item = data_row[8]
        self.item_details = data_row[9]
        self.gp = data_row[10]
        self.gp_address = data_row[11]

    def entry_summary(self):
        return [
            self.date,
            self.patient,
            self.address,
            self.dob,
            self.ppsn,
            self.gender,
            self.item,
            self.item_details,
        ]


def get_outlook():
    try:
        outlook = win32.Dispatch("outlook.application")
        return outlook
    except:
        # Exit the program if Outlook is not open
        print("Error: Outlook is not open - emails cannot be created.")
        sys.exit(1)


def get_data(file):
    # obtain data from our file
    data = []
    with open(file) as f:
        csv_reader = csv.reader(f, delimiter=",")
        for row in csv_reader:
            amended_row = process_row(row)
            data.append(PatientEntry(amended_row))
    return data[1:]


def process_row(row):  # TODO
    """ 
    ['Date Dispensed', 'Patient Name', 'Street', 'Town or City', 'Birth Date',
    'PPSN No', 'Gender', 'Qty', 'Script Dispensed As', 'Directions Expanded',
    'Contract GP Name', 'Contract GP Address']
    """
    return row


def create_email(account, address, subject, body):
    # creates an email
    outlook = get_outlook()
    mail = outlook.CreateItem(0)
    mail.To = address
    mail.Subject = subject
    mail.Body = body
    # set the "send from" account using arcane methods
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
    # save a draft
    mail.Save()
    # open the email in a popup window (disabled in favour of saving a draft)
    # mail.Display(False)
    return mail


def format_body(name, patients):
    # composes the message text
    # TODO: rich text version instead to allow for formatting?
    greeting = "Dear {},\n".format(name)
    general_body = "For your information, the below patients of yours were recently vaccinated in our pharmacy. Vaccine details, including batch and expiry, are below.\n"
    patient_details = "{}".format(patients)
    sign_off = "\nKind regards,\n"
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


def create_recipient_list(data):  # TODO
    # create our list of recipients based on the data we have
    recipient_list = []
    for entry in data:
        gp = entry.gp
        entry_summary = entry.entry_summary()
        if gp in recipient_list:
            recipient_list[recipient_list.index(gp)].add_patient_entry(entry_summary)
        else:
            recipient_list.append(
                Recipient(entry.gp, entry.gp_address, "", entry_summary)
            )

    """ recipient_list = {
        "Adrian Test": ["Patient A", "Patient B"],
        "Beatrice Test": ["Patient C", "Patient D"],
    }  # debug values """
    return recipient_list


def create_emails(account, recipient_list):
    # for each recipient in list, create & save a draft email
    for recipient in recipient_list:
        name = recipient.name
        patients = recipient.patients
        address = recipient.address
        email = recipient.email
        print("Composing email for {}".format(name))
        body = format_body(name, patients)
        # print(body)  # debug
        subject = "Vaccine Report - {}".format(name)
        # mail = create_email(account, email, subject, body) # debug - disabled while testing


def main():
    # find the file with data
    data = get_data("Flu Vacc Report.csv")
    # generate list of recipients
    recipient_list = create_recipient_list(data)
    # select correct account
    account = ""  # select_account("healthmail")
    # create emails
    create_emails(account, recipient_list)


if __name__ == "__main__":
    main()
