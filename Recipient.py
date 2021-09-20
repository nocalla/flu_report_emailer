class Recipient:
    def __init__(self, name, address, email, patient_entry):
        self.name = name
        self.address = address
        self.email = email
        self.patients = [
            [
                "Date",
                "Patient Name",
                "Patient Address",
                "Date of Birth",
                "PPSN",
                "Gender",
                "Vaccine Administered",
                "Batch No. & Expiry",
            ]
        ]
        self.patients.append(patient_entry)
        self.surname = name.split()[-1]

    def add_patient_entry(self, patient_entry):
        self.patients.append(patient_entry)

    def generate_patient_summary(self):
        return "\n".join(self.patients)
