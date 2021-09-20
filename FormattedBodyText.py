class FormattedBodyText:
    def __init__(self, template: str) -> None:
        self.template = template

    def format_body(self, name: str, patients: list) -> str:
        """composes the message text by filling in the specified html template

        Args:
            name (str): recipient name
            patients (list): list of individual entries to tabulate

        Returns:
            str: text of email body to send
        """

        with open(self.template, "r") as f:
            html_template = f.read()
        with open("test_email.html", "w") as f:  # TODO: remove debug file creation
            new_html = html_template.format(name, self.html_table(patients))
            f.write(new_html)

        return new_html

    def html_table(self, data: list) -> str:
        """returns a html table derived from a list of data rows

        Args:
            data (list): list of rows (each row is in list form) with first row being headings

        Returns:
            str: html tablified version of input list
        """
        table = "<table>"
        headers = "</th><th>".join(data[0])
        table += "<tr><th>{}</th></tr>\n".format(headers)
        for row in data[1:]:
            cells = "</td><td>".join(row)
            table += "<tr><td>{}</td></tr>\n".format(cells)

        table += "</table>"
        return table
