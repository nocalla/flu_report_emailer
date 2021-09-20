import csv
from PatientEntry import PatientEntry
from Recipient import Recipient
from EmailHandler import EmailHandler
from FormattedBodyText import FormattedBodyText


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
        surname = recipient.surname
        patients = recipient.patients
        email = recipient.email
        print("Composing email for {} ({} patients)".format(name, len(patients)))
        body_object = FormattedBodyText("email_template.html")
        body = body_object.format_body(surname, patients)
        subject = "Vaccine Report - {}".format(name)
        email_list.append([email, subject, body])
    return email_list


def email_list_iterate(account_str: str, email_list: list):
    """trigger email creation for each item in our list

    Args:
        account_str (str): term to search Outlook account for when sending email
        email_list (list): list of individual email details
    """
    for mail_details in email_list:
        email_object = EmailHandler(account_str)
        email_object.create_email(mail_details)


def main():
    # find the file with data
    data = get_data("Flu Vacc Report.csv")
    # get existing email contacts
    email_addresses = get_email_addresses("Healthmail Addresses.csv")
    # generate list of recipients
    recipient_dict = create_recipient_dict(data, email_addresses)
    # create email details for all recipients
    email_list = compose_email_details(recipient_dict)
    # send all emails from the specified account
    email_list_iterate("healthmail", email_list)


if __name__ == "__main__":
    main()
