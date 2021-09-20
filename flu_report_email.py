import csv
import sys

import win32com.client as win32


class Recipient:
    def __init__(self, name, address, email, patient_entry):
        self.name = name
        self.address = address
        self.email = email
        self.patients = [
            "Date; Patient Name; Patient Address; Date of Birth; PPSN; Gender; Vaccine Administered; Batch No. & Expiry"
        ]
        self.patients.append(patient_entry)

    def add_patient_entry(self, patient_entry):
        self.patients.append(patient_entry)

    def generate_patient_summary(self):
        return "\n".join(self.patients)


class PatientEntry:
    def __init__(self, heading_indices, data_row):
        self.date = data_row[heading_indices["Date Dispensed"]]
        self.name = data_row[heading_indices["Patient Name"]]
        self.address = "{}, {}".format(
            data_row[heading_indices["Street"]],
            data_row[heading_indices["Town or City"]],
        )
        self.dob = data_row[heading_indices["Birth Date"]]
        self.ppsn = data_row[heading_indices["PPSN No"]]
        self.gender = data_row[heading_indices["Gender"]]
        self.item = data_row[heading_indices["Script Dispensed As"]]
        self.item_details = data_row[heading_indices["Directions Expanded"]]
        self.gp = data_row[heading_indices["Contract GP Name"]]
        self.gp_address = data_row[heading_indices["Contract GP Address"]]
        self.consent = self.check_consent()

    def check_consent(self):
        if "No GP" in self.item_details:
            return False
        return True

    def entry_summary(self):
        summary_list = [
            self.date,
            self.name,
            self.address,
            self.dob,
            self.ppsn,
            self.gender,
            self.item,
            self.item_details,
        ]
        return "; ".join(summary_list)


def get_outlook():
    try:
        outlook = win32.Dispatch("outlook.application")
        return outlook
    except:
        # Exit the program if Outlook is not open
        print("Error: Outlook is not open - emails cannot be created.")
        sys.exit(1)


def get_data(file: str) -> list:
    """extract our data from our source file

    Args:
        file (str): filename

    Returns:
        list: list of PatientEntry objects
    """
    data = []
    with open(file) as f:
        csv_reader = csv.reader(f, delimiter=",")
        # get the index for each heading (allow for changing order)
        # NB: the next function moves the CSV reader on 1 row, so don't need to allow for heading skipping in loop
        heading_indices = get_heading_indices(next(csv_reader))
        for row in csv_reader:
            patient_entry = PatientEntry(heading_indices, row)
            if patient_entry.consent:
                data.append(patient_entry)
    return data


def get_email_addresses(file: str) -> dict:
    """get a dictionary mapping recipient names to email addresses

    Args:
        file (str): filename of our CSV file with email addresses

    Returns:
        dict: dictionary mapping recipient names to email addresses
    """
    data = dict()
    with open(file) as f:
        csv_reader = csv.reader(f, delimiter=",")
        next(csv_reader)
        for row in csv_reader:
            data[row[0]] = row[1]
    return data


def get_heading_indices(row: list) -> dict:
    """generates a dictionary mapping desired headings to row indices to allow for changing order of columns in source data

    Args:
        row (list): row of data from CSV file

    Returns:
        dict: dictionary of heading matched with row index
    """
    headings = [
        "Date Dispensed",
        "Patient Name",
        "Street",
        "Town or City",
        "Birth Date",
        "PPSN No",
        "Gender",
        "Qty",
        "Script Dispensed As",
        "Directions Expanded",
        "Contract GP Name",
        "Contract GP Address",
    ]

    heading_indices = dict()
    for heading in headings:
        heading_indices[heading] = row.index(heading)
    return heading_indices


def create_email(account: object, mail_details: list) -> object:
    """create an Outlook email object based on mail_details list and save as draft under specified Outlook account

    Args:
        account (object): specified Outlook account
        mail_details (list): email details in form [to_address, subject, body_text]

    Returns:
        object: outlook mail object
    """
    outlook = get_outlook()
    mail = outlook.CreateItem(0)
    mail.To = mail_details[0]
    mail.Subject = mail_details[1]
    mail.Body = mail_details[2]
    # set the "send from" account using arcane methods
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
    # save a draft
    mail.Save()
    # open the email in a popup window (disabled in favour of saving a draft)
    # mail.Display(False)
    return mail


def format_body(
    name: str, patients: list
) -> str:  # TODO: version to allow for formatting? maybe use a template that we can change out values in

    """composes the message text

    Args:
        name (str): recipient name
        patients (list): list of individual entries to tabulate

    Returns:
        str: text of email body to send
    """

    greeting = "Dear {},\n".format(name)
    general_body = "For your information, the below patients of yours were recently vaccinated in our pharmacy.\nVaccine details, including batch and expiry, are listed below.\n"
    patient_details = "{}".format(patients)
    sign_off = "\nKind regards,\n"
    body_string = greeting + general_body + patient_details + sign_off

    return body_string


def select_account(search: str) -> object:
    """select the account to send the email from

    Args:
        search (str): desired search term for account name

    Returns:
        object: outlook account object for the desired account
    """
    outlook = get_outlook()
    accounts = outlook.Session.Accounts
    for account in accounts:
        if search in str(account):
            from_account = account
            break
    print("Account selected: {}".format(from_account))
    return from_account


def create_recipient_dict(data: list, email_addresses: dict) -> dict:
    """create our dictionary of recipients based on the data we have. use existing email address data to populate each recipient if available

    Args:
        data (list): list of PMR entries    

    Returns:
        dict: dictionary of recipients and their matching PMR entries
    """

    recipient_dict = dict()
    for entry in data:
        if entry.consent:
            recipient = entry.gp
            entry_summary = entry.entry_summary()

            if recipient in recipient_dict:
                # if our recipient is in the dictionary, add the entry to their section
                recipient_dict[recipient].add_patient_entry(entry_summary)
            else:
                # if our recipient is not in the dictionary, add them to the dictionary
                # check if we have an email address saved for the recipient already
                email = ""
                if recipient in email_addresses:
                    email = email_addresses[recipient]
                recipient_dict[recipient] = Recipient(
                    entry.gp, entry.gp_address, email, entry_summary,
                )

    return recipient_dict


def compose_email_details(
    recipient_dict: dict,
) -> list:  # TODO: remove debug print output
    """for each recipient in list, create email details

    Args:
        recipient_dict (dict): dictionary of recipients and their matching PMR entries

    Returns:
        list: list of lists of email details
    """
    email_list = []
    for key in recipient_dict:
        recipient = recipient_dict[key]
        name = recipient.name
        patients = recipient.generate_patient_summary()
        address = recipient.address
        email = recipient.email
        print("Composing email for {}".format(name))
        body = format_body(name, patients)
        subject = "Vaccine Report - {}".format(name)
        email_list.append([email, subject, body])
        print(
            "To: {}\nSubject: {}\n{}".format(email, subject, body)
        )  # DEBUG - display email details in terminal
    return email_list


def email_list_iterate(account: object, email_list: list):
    """trigger email creation for each item in our list

    Args:
        account (object): outlook account to use for sending
        email_list (list): list of individual email details
    """
    for mail_details in email_list:
        email_object = create_email(account, mail_details)


def main():
    # find the file with data
    data = get_data("Flu Vacc Report.csv")
    # get existing email contacts
    email_addresses = get_email_addresses("Healthmail Addresses.csv")
    # generate list of recipients
    recipient_dict = create_recipient_dict(data, email_addresses)
    # create email details
    email_list = compose_email_details(recipient_dict)
    # select correct account
    account = select_account("healthmail")
    # send emails
    email_list_iterate(account, email_list)


if __name__ == "__main__":
    main()
