from openpyxl import load_workbook
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2pdf import convert
from RPA.Email.ImapSmtp import ImapSmtp
from dotenv import load_dotenv

class Excel:
    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = load_workbook(self.file_path)
        self.sheet = self.wb.active
    
    def get_total_rows(self):
        return self.sheet.max_row - 1
    
    def get_headers(self):
        if self.get_total_rows() > 0:
            return self.read_from_file(1)

    def read_from_file(self, row_to_read):
        for row in self.sheet.iter_rows(min_row=row_to_read, max_row=row_to_read, values_only=True):
            return row

class Invoice:
    def __init__(self, invoice_template, invoice_folder):
        self.invoice_template = os.path.join(".", invoice_template)
        self.invoice_folder = os.path.join(".", invoice_folder)

    def generate_invoices(self, data):
        template = Document(self.invoice_template)
        for p in template.paragraphs:
            for key, value in data.items():
                if f"[{key}]" in p.text:
                    p.text = p.text.replace(f"[{key}]", value)

        file_name = os.path.join(self.invoice_folder, data["invoice_number"])
        docx_path = f"{file_name}.docx"
        pdf_path = f"{file_name}.pdf"

        template.save(docx_path)
        convert(docx_path, pdf_path)
        os.remove(docx_path)

        return pdf_path
    
    def read_template(self, data):
        template = Document(self.invoice_template)

        for p in template.paragraphs:
            full_text = "".join(run.text for run in p.runs)
            new_text = full_text
            for key, value in data.items():
                new_text = new_text.replace(f"[{key}]", str(value))
            if new_text != full_text:
                p.clear()
                p.add_run(new_text)

        for table in template.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        full_text = "".join(run.text for run in para.runs)
                        new_text = full_text
                        for key, value in data.items():
                            new_text = new_text.replace(f"[{key}]", str(value))
                        if new_text != full_text:
                            para.clear()
                            para.add_run(new_text)
        
        for table in template.tables:
            if len(table.columns) == 4 and table.cell(0, 0).text.strip().upper() == "DESCRIPTION":
                if len(table.rows) > 1:
                    table.rows[1].cells[0].text = data["product"]
                    table.rows[1].cells[1].text = str(data["quantity"])
                    table.rows[1].cells[2].text = str(data["rate"])
                    table.rows[1].cells[3].text = str(data["subtotal"])
                break
        
        for table in template.tables:
            for row in table.rows:
                for i, cell in enumerate(row.cells):
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip() == "SUBTOTAL":
                            next_cell = row.cells[-1]
                            next_cell.text = str(data["subtotal"])
                        elif paragraph.text.strip() == "TAX":
                            next_cell = row.cells[-1]
                            next_cell.text = str(data["total_tax"])
                        elif paragraph.text.strip() == "GRAND TOTAL":
                            next_cell = row.cells[-1]
                            next_cell.text = str(data["total"])

            
        # Save the edited document
        file_name = os.path.join(self.invoice_folder, data["invoice_number"])
        docx_path = f"{file_name}.docx"
        pdf_path = f"{file_name}.pdf"

        template.save(docx_path)
        convert(docx_path, pdf_path)
        os.remove(docx_path)
        return pdf_path

    
class DataFormat:
    def __init__(self, tax_percent, data_headers):
        self.tax = tax_percent / 100
        #need to add code/ function to get data dictionary acc to headers input by the user

    def format_data(self, details):
        data = {
            "invoice_number": details[0],
            "client_name": details[1],
            "client_email": details[2],
            "invoice_date": details[3].strftime('%d-%m-%Y'),
            "due_date": details[4].strftime('%d-%m-%Y'),
            "product": details[5],
            "quantity": str(details[6]),
            "rate": details[7],
            "payment_mode": details[8],
            "billing_address": details[9],
            "shipping_address": details[10]
        }

        subtotal = int(details[6]) * int(details[7])
        total_tax = subtotal * self.tax
        total = subtotal + total_tax

        data["subtotal"] = str(subtotal)
        data["total_tax"] = str(total_tax)
        data["total"] = str(total)
        return data

class Email:
    def __init__(self):
        self.mail = ImapSmtp()
        self.gmail_account = os.getenv("GMAIL_ACCOUNT")
        self.gmail_password = os.getenv("GMAIL_PASSWORD")

        self.mail.authorize(
            account=self.gmail_account,
            password=self.gmail_password,
            smtp_server="smtp.gmail.com",
            smtp_port=587,
        )

    def send_mail(self, file_path, invoice_num, client_name, client_email=None):
        if not client_email:
            client_email = self.gmail_account
        if not client_email:
            raise ValueError("No valid recipient email provided.")

        body = f"Dear {client_name}, \nPlease find attached the invoice {invoice_num} for the recent purchase/service. \nIf you have any questions or concerns, feel free to contact us.\nThank you for your business!\nBest regards,\nAnushka\n"

        self.mail.send_message(
            recipients=client_email,
            cc=self.gmail_account,
            sender=self.gmail_account,
            subject=f"Invoice {invoice_num} from Anushka",
            body=body,
            attachments=file_path
        )

class InvoiceGenerator:
    def __init__(self, file_name, tax_percent, invoice_template, invoice_folder):
        self.file_path = os.path.join(".", file_name)
        self.tax = tax_percent
        self.invoice_template = invoice_template
        self.invoice_folder = invoice_folder

    def main(self):
        load_dotenv()
        excel_work = Excel(self.file_path)
        headers = excel_work.get_headers()
        
        invoice_generator = Invoice(self.invoice_template, self.invoice_folder)
        email = Email()
        
        data_formatter = DataFormat(self.tax, headers)

        total_invoices = excel_work.get_total_rows()

        for i in range(2, total_invoices + 2):
            details = excel_work.read_from_file(i)
            data = data_formatter.format_data(details)
            file_path = invoice_generator.read_template(data)
            email.send_mail(file_path, data["invoice_number"], data["client_name"], data["client_email"])

if __name__ == "__main__":
    invoice = InvoiceGenerator("invoice_details.xlsx", 10, "invoice-basic.docx", "invoices")
    invoice.main()