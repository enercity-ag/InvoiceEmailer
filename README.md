# **Introduction** #

The class InvoiceEmailer automates the process of reading invoice data from Excel files, matching departments to responsible employees, generating email addresses, and sending reminder emails via Outlook.

In the actual version, the program reads an excel file which consists of two sheets. One for basware, one for workday. 
It reads the open tasks (i.e. invoices that still have to be checked etc) in the basware file and finds the corresponding person(s) in charge from the worday file.
An email containing a reminder about the open tasks is then sent to them. 
The code serves as a basis for future improvements. 
