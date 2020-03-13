#
# main
#
# Excelデータ読み込み
import sys
import win32com.client
from pathlib import Path
from invoice import Invoice, InvoiceDetail
import invoice_util

PDF_DIR_PATH = "../pdf"

try:
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    abs_path = str(Path(r"../invoice.xlsm").resolve())
    wb = app.Workbooks.Open(abs_path)
except:
    print('can\'t open invoice file')
    sys.exit(-1)

wb.Activate()
ws = wb.WorkSheets("請求データ")
iv_data = invoice_util.get_excel_data(ws)

# 請求書データ生成
invoices = [] 
for row in iv_data:
    id = row[1]
    exist_invoice = invoice_util.get_invoice_from_list(invoices, id)
    if not exist_invoice:
        invoice_data = Invoice(row)
        invoice_data.add_detail(InvoiceDetail(row))
        invoices.append(invoice_data)
    else:
        exist_invoice.add_detail(InvoiceDetail(row))

invoice_doc = wb.WorkSheets("請求書テンプレート")
for iv in invoices:
    invoice_util.create_invoice_pdf(app, wb, invoice_doc, iv, PDF_DIR_PATH)

wb.Close()
app.Quit()