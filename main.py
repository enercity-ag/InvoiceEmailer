from invoice_emailer import InvoiceEmailer

if __name__ == "__main__":
    emailer = InvoiceEmailer('Y:/BaswareReminder/InvoiceEmailer/test_data/your_file.xlsx', 'Y:/Baswarereminder/InvoiceEmailer/results/your_file_with_emails.xlsx')
    emailer.load_data()
    emailer.dict_emails_oe()
    emailer.add_emails_to_basware()
    emailer.save_new_file()
    emailer.send_emails()

