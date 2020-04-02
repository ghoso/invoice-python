#
# main
#
import sys
import win32com.client
from pathlib import Path
from invoice import Invoice, InvoiceDetail
import invoice_util

# PDF出力ディレクトリ
PDF_DIR_PATH = "../pdf"

try:
    # Excel呼び出し
    app = win32com.client.Dispatch("Excel.Application")
    # Excel非表示
    app.Visible = False
    # Excelファイルパス
    abs_path = str(Path(r"../invoice.xlsm").resolve())
    # Excelファイルオープン
    wb = app.Workbooks.Open(abs_path)
except:
    print('can\'t open invoice file')
    sys.exit(-1)

wb.Activate()
# ワークシート
ws = wb.WorkSheets("請求データ")
# Excelデータ取得
iv_data = invoice_util.get_excel_data(ws)

# 請求書データ
invoices = [] 
# Excelデータから請求書データ生成
for row in iv_data:
    id = row[1]
    exist_invoice = invoice_util.get_invoice_from_list(invoices, id)
    # 初めての請求書番号のデータだったら、Invoiceデータを作成
    if not exist_invoice:
        invoice_data = Invoice(row)
        invoice_data.add_detail(InvoiceDetail(row))
        invoices.append(invoice_data)
    else:
        # 既に取得済みの請求書番号のデータ
        exist_invoice.add_detail(InvoiceDetail(row))

# 請求書雛形読み込み
invoice_doc = wb.WorkSheets("請求書テンプレート")
# 請求書生成
for iv in invoices:
    invoice_util.create_invoice_pdf(app, wb, invoice_doc, iv, PDF_DIR_PATH)

wb.Close()
app.Quit()