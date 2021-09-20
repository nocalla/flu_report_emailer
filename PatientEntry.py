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
        return summary_list
