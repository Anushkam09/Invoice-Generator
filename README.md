# 🧾 Invoice Generator using Python (RPA + Word + Excel + Email)

This project automatically generates professional invoices using data from an Excel sheet, fills them into a Word template, converts them to PDF, and emails them to clients. It uses Python libraries like `openpyxl`, `python-docx`, `docx2pdf`, and `RPA.Email.ImapSmtp` to create a seamless automation workflow.

---

## ✅ Features Implemented

- Read invoice data from Excel
- Fill placeholders in a Word template with invoice details
- Email the generated invoice to the client
- Allows multiple products per invoice
- Uses professional looking template

---

## 🛠️ Setup Instructions

### 1. Clone the repository

```
git clone https://github.com/yourusername/invoice-generator.git
cd invoice-generator
```

### 2. Install required packages

```bash
pip install -r requirements.txt
```

### 3. Prepare `.env` file

Create a `.env` file in the root directory and add:

```
GMAIL_ACCOUNT=your_email@gmail.com
GMAIL_PASSWORD=your_app_password
```

> 💡 Use a Gmail **App Password** if you have 2FA enabled.

---

## 📂 Folder Structure

```
invoice-generator/
│
├── invoice_details.xlsx         # Excel with invoice data
├── invoice-basic.docx
├── invoices/                    # Folder to save generated PDFs
├── .env                         # Email credentials
├── requirements.txt             # Dependencies
└── invoice_generator.py         # Main script
```

---

## 🔄 How to Run

```bash
python invoice_generator.py
```

Invoices will be generated and emailed one by one based on the Excel file rows.

---

## 🚧 Work in Progress

* ⏳ Allow personalized templates
* ⏳ Add a web interface using Flask or Streamlit
* ⏳ Add logs, error handling, and email status tracking
* ⏳ Export invoice history to a separate Excel/CSV file

---

Made with 💻 by Anushka Mahajan 🐈

