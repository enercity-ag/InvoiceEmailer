import pandas as pd # For reading and manipulating excel files
import collections # used here for defaultdict to group messages by email
import win32com.client as win32 # Allows interaction wih Micorosoft Oulook to send emails.


class InvoiceEmailer:

    def __init__(self, input_file, output_file, domain = 'enercity.de'):
        self.input_file = input_file
        self.output_file = output_file
        self.domain = domain
        self.sheet_basware = None
        self.sheet_workday = None
        self.email_dict = {}
   
    def load_data(self):
        """ Read the excel file, which consists of two sheets,
        one for basware, one for workday.
        """
        self.sheet_basware = pd.read_excel(self.input_file, sheet_name = 0)
        self.sheet_workday = pd.read_excel(self.input_file, sheet_name = 1)
 
    def extract_prefix(self, oe):
        """ Extract the prefix before the dash. For example, F-RN returns F.
        This identifies the person(s) in charge for the given department.
        """
        if pd.isna(oe):
            return ''
        return str(oe).split('-')[0]
   
    def format_email(self, name):
        """ Takes Surname Name and returns the corresponding email address name.surname@domain
        """
        if pd.isna(name) or not isinstance(name, str):
            return ''
        parts = name.strip().lower().split()
        if len(parts) < 2:
            return ''
        first, last = parts[-1], parts[0]
        return f"{first}.{last}@{self.domain}"
   
    def dict_emails_oe(self):
        """ Dictionary between prefixes and emails. """
        for _, row in self.sheet_workday.iterrows():
            oe_prefix =str(row.get('OE','')).strip()
            if '-' not in oe_prefix and oe_prefix.isalpha(): # Only accept OEs without dashes and with letters
                name = row.get('Name Gesamt', '')
                email = self.format_email(name)
                if email:
                    self.email_dict.setdefault(oe_prefix, []).append(email)

    def add_emails_to_basware(self):
        """ Adds an additional column in the basware file, which contain the email(s)
        of the person(s) in charge of the department. """
        self.sheet_basware['OE_prefix'] = self.sheet_basware['OE'].apply(self.extract_prefix)
        self.sheet_basware['Emails'] = self.sheet_basware['OE_prefix'].apply(lambda prefix: '; '.join(self.email_dict.get(prefix, [])))
       
    def save_new_file(self):
        """ Create a separate basware file with the emails column
        """
        self.sheet_basware.to_excel(self.output_file, index = False)
        print(f"Updated file saved as: {self.output_file}")

    def send_emails(self):
        """ Send emails to the people in charge (Vorgesetzten) with a customized message. """
        # Create Outlook interface
        outlook = win32.Dispatch('outlook.application')

        # Read the sheet of the new file
        self.sheet_basware = pd.read_excel(self.output_file, sheet_name = 0)

        # Dictionary to collect messages per OE group
        oe_messages = collections.defaultdict(list)

        # Read emails, invoice numbers and delate days
        for _, row in self.sheet_basware.iterrows():
            oe_prefix = row.get('OE_prefix', '')
            invoice_number = row.get('Rechnungsnummer', '')
            days_due = row.get('Ausstehend seit', '')
            supplier = row.get('Lieferantencode', '')

            # Skip invalid rows
            if pd.isna(oe_prefix) or pd.isna(invoice_number) or pd.isna(days_due):
                 continue
           
            message = f"Die Rechnung {invoice_number} mit Lieferantencode {supplier} ist seit {days_due} Tagen ausstehend. "
            oe_messages[oe_prefix].append(message)
           
        # Send one email per OE group. Modify the introductory message as you wish
        for oe_prefix, messages in oe_messages.items():
            recipients = self.email_dict.get(oe_prefix, [])
            if not recipients:
                continue
           
            recipient_str = ';'.join(recipients)
            subject = "Quartalsabschluss // Bearbeitung ausstehender Aufgaben im Basware"
            body =  (
            "Liebe Kolleg*innen,\n\n"
            "im Sinne eines schnelleren Prozessdurchlaufes und der stringenteren Einhaltung der Abschluss-Terminpläne, "
            "widmen wir dem Thema der offenen Aufgaben bei Eingangsrechnungen-Bearbeitungen derzeit bekanntlich mehr Augenmerk.\n"
            "Bitte die folgenden offenen Basware - Aufgaben für den Verantwortungsbereich beachten.\n\n"
            )
            body += '\n'.join(messages)
            body +=  (
            "\n\nAufgrund der bislang sehr manuellen Arbeiten sind wir derzeit noch in der Fortentwicklung"
            "einer automatisierten Auswertung und Zuordnung. Daher steht die Namens-/Bereichszuordnung noch unter Vorbehalt."
            "Sollten euch fehlerhafte Zuordnungen auffallen oder ihr habt weiterbringende Hinweise, meldet euch gerne."
            "Vielen Dank für eure Unterstützung, Bearbeitung und Koordination.\n"
            "Positive Grüße\n\n Das Team Rechnungswesen Nebenbücher"
            )

            # Create email. One for each recipient
            mail = outlook.CreateItem(0)
            mail.To = recipient_str
            mail.CC = "susanne.urban@enercity.de" #Add CC recipients
            mail.Subject = subject
            mail.Body = body
            mail.Display() # Check before sending
            # mail.send() # Finally send




    



