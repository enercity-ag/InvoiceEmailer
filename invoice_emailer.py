import pandas as pd
import win32com.client as win32

class InvoiceEmailer:

    def __init__(self, input_file, output_file, domain = 'enercity.de'):
        self.input_file = input_file
        self.output_file = output_file
        self.domain = domain
        self.sheet_basware = None
        self.sheet_workday = None
        self.dict = {}
    
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
            oe_prefix = self.extract_prefix(row.get('OE', ''))
            name = row.get('Name Gesamt', '')
            email = self.format_email(name)
            if email: 
                self.dict.setdefault(oe_prefix, []).append(email)

    def add_emails_to_basware(self):
        """ Adds an additional column in the basware file, which contain the email(s) 
        of the person(s) in charge of the department. """
        self.sheet_basware['OE_prefix'] = self.sheet_basware['OE'].apply(self.extract_prefix)
        self.sheet_basware['Emails'] = self.sheet_basware['OE_prefix'].apply(lambda prefix: '; '.join(self.dict.get(prefix)))
        self.sheet_basware.drop(columns = ['OE_prefix'], inplace = True)

    def save_new_file(self):
        """ Create a separate basware file with the emails columsn
        """
        self.sheet_basware.to_excel(self.output_file, index = False)
        print(f"Updated file saved as: {self.output_file}")

    def send_emails(self):
        """ Send emails with a customized message. """
        # Create Outlook interface
        outlook = win32.Dispatch('outlook.application')

        # Read the sheet of the new file
        self.sheet_basware = pd.read_excel(self.output_file, sheet_name = 0)

        # Read emails, invoice numbers and delate days
        for _, row in self.sheet_basware.iterrows():
            email_list = row.get('Emails', '')
            invoice_number = row.get('Rechnungsnummer', '')
            days_due = row.get('Ausstehend seit', '')

            # Skip invalid rows
            if pd.isna(email_list) or pd.isna(invoice_number) or pd.isna(days_due):
                 continue

            # Clean up email list
            recipients = [email.strip() for email in email_list.split(';') if email.strip()]
            if not recipients:
                continue

            # Compose message (modify it as you wish)
            subject = f"Ihre Rechnung {invoice_number} ist überfällig"
            body = f"Guten Tag, bitte beachten Sie: die Rechnung {invoice_number} ist seit {days_due} Tagen ausstehend. "

            # Create email
            mail = outlook.CreateItem(0)
            mail.To = '; '.join(recipients)
            mail.Subject = subject
            mail.Body = body
            mail.Display() # Check before sending
            # mail.send() # Finally send





