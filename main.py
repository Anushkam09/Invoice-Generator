from openpyxl import load_workbook
import os
from datetime import datetime
from docx import Document
from docx2pdf import convert
from RPA.Email.ImapSmtp import ImapSmtp
from dotenv import load_dotenv

class Excel:
    def __init__(self, file_path):
        self.file_path = file_path
    
    def get_total_rows(self):
        wb = load_workbook(self.file_path)
        sheet = wb.active
        return sheet.max_row -1

    def read_from_file(self,row_to_read):
        wb = load_workbook(self.file_path)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=row_to_read, max_row=row_to_read, values_only=True):
            return row

class Invoice:
    def __init__(self, invoice_template, invoice_folder):
        self.invoice_template = f"./{invoice_template}"
        self.invoice_folder = f"./{invoice_folder}"
    
    def generate_invoices(self, data):
        template = Document(self.invoice_template)
        for p in template.paragraphs:
            for key, value in data.items():
                if key in p.text:
                    p.text = p.text.replace(f"[{key}]", value)
                    print(p.text)
        

        file_name = f"{self.invoice_folder}/{data["invoice_number"]}"
        template.save(f"{self.invoice_folder}/{data["invoice_number"]}.docx")
        convert(f"{file_name}.docx", f"{file_name}.pdf" )
        os.remove(f"{file_name}.docx")
        return f"{file_name}.pdf", data["invoice_number"], data["client_name"], None
        
        
        # for table in template.tables:
        #     for row in table.rows:
        #         for cell in row.cells:
        #             for p in cell.paragraphs:
        #                 print(p.text)

class DataFormat:
    def __init__(self, tax_percent):
        self.tax = tax_percent/100
    
    def format_data(self, details):
        data = {}
        data["invoice_number"] = details[0]
        data["client_name"] = details[1]
        data["client_email"] = details[2]
        data["invoice_date"] = details[3].strftime('%d-%m-%Y')
        data["due_date"] = details[4].strftime('%d-%m-%Y')
        data["product"] = details[5]
        data["quantity"] = details[6]
        data["rate"] = details[7]
        data["payment_mode"] = details[8]
        data["billing_address"] = details[9]
        data["shipping_address"] = details[10]

        data["subtotal"] = int(data["quantity"]) * int(data["rate"])
        data["total_tax"] = data["subtotal"]*self.tax
        data["total"] = data["subtotal"] + data["total_tax"]

        data["quantity"] = str(data["quantity"])
        data["subtotal"] = str(data["subtotal"])
        data["total_tax"] = str(data["total_tax"])
        data["total"] = str(data["total"])
        return data

class Email:
    def __init__(self):
        self.mail = ImapSmtp()

        self.gmail_account =  os.getenv("GMAIL_ACCOUNT")
        self.gmail_password = os.getenv("GMAIL_PASSWORD")

        self.mail.authorize(
            account=self.gmail_account,
            password=self.gmail_password,
            smtp_server="smtp.gmail.com",
            smtp_port=587,
        )
    
    def send_mail(self,file_path, invoice_num,client_name, client_email = None):
        if not client_email:
            client_email = self.gmail_account
        if not client_email:
            raise ValueError("No valid recipient email provided.")
        data = f"Dear {client_name}, \nHere is your invoice from Anushka"

        print(client_email)
        self.mail.send_message(
            recipients=client_email,
            sender= self.gmail_account,
            subject= f"{invoice_num}: Invoice from Anushka",
            body=data,
            attachments=file_path
        )
        return

class BillTemplate:
    def __init__(self, tax_percent):
        self.tax = tax_percent/100

    def create_bill(self, details):
        invoice_number = details[0]
        client_name = details[1]
        client_email = details[2]
        invoice_date = details[3]
        due_date = details[4]
        product = details[5]
        quantity = details[6]
        rate = details[7]
        payment_mode = details[8]
        billing_address = details[9]
        shipping_address = details[10]

        subtotal = int(quantity) * int(rate)
        total_tax = subtotal*self.tax
        total = subtotal + total_tax

        print(f"""
-----------------------------------
Invoice Number:    {invoice_number}
Invoice Date:      {invoice_date.strftime('%d-%m-%Y')}
Client Name:       {client_name}
Client Email:      {client_email}
Billing Address:   {billing_address}
Shipping Address:  {shipping_address}

-----------------------------------
Product \t Quantity \t Rate
-----------------------------------
{product} \t {quantity} \t\t {rate}

-----------------------------------
-----------------------------------
Subtotal: \t {subtotal}
GST ({self.tax*100}%) \t {total_tax}
-----------------------------------
Grand Total: \t {total}

-----------------------------------
Due Date: \t {due_date.strftime('%d-%m-%Y')}
Payment Mode: \t {payment_mode}
-----------------------------------""")

class InvoiceGenerator:
    def __init__(self, file_name, tax_percent, invoice_template, invoice_folder):
        self.file_path = f"./{file_name}"
        self.tax = tax_percent
        self.invoice_template = invoice_template
        self.invoice_folder = invoice_folder
    
    def main(self):
        load_dotenv()
        excel_work = Excel(self.file_path)
        make_bills = BillTemplate(self.tax)
        invoice_generator = Invoice(self.invoice_template, self.invoice_folder)
        data_formatter = DataFormat()
        email = Email()
        total_invoices = excel_work.get_total_rows()
        #row1: header, start from row 2
        for i in range(2, total_invoices + 2):
            details = excel_work.read_from_file(i)
            # make_bills.create_bill(details)
            data = data_formatter.format_data(details)
            file_path, invoice_num, name, mail_id = invoice_generator.generate_invoices(data)
            email.send_mail(file_path, invoice_num, name, mail_id)
            

if __name__ == "__main__":
    invoice = InvoiceGenerator("invoice_details.xlsx", 10, "invoice-practice.docx", "invoices")
    invoice.main()