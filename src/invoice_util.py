#
# invoice_util.py
#
import datetime
import pathlib

# 定数
INVOICE_HEADER_ROW = 3
INVOICE_COL = 15
COMPANY_NAME_ROW = 3
COMPANY_NAME_COL = 1
PERSONNEL_ROW = 4
PERSONNEL_COL = 4
INVOICE_ID_ROW = 3
INVOICE_ID_COL = 14
TODAY_ROW = 4
TODAY_COL = 14
TITLE_ROW = 6
TITLE_COL = 3
DUE_DATE_ROW = 15
DUE_DATE_COL = 13
DETAIL_START_ROW = 18
DETAIL_NUM_COL = 1
PRODUCT_NAME_COL = 2
UNIT_NUM_COL = 10
UNIT_PRICE_COL = 12
DESCRIPTION_ROW = 36
DESCRIPTION_COL = 3

# datetimeデータから文字列出力
def pdate_to_str(pdate):
    year = pdate.strftime("%Y年")
    month = pdate.strftime("%m月").lstrip("0")
    day = pdate.strftime("%d日").lstrip("0")
    return(year+month+day)

# Excelの日付データから文字列変換
def excel_date_to_str(excel_date):
    pdate = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=excel_date - 2)
    pdate_str = pdate_to_str(pdate)
    return(pdate_str)

# 請求書データをエクセルシートから取得 
def get_excel_data(ws):
    invoice_data = [[0]]
    nrows = ws.UsedRange.Rows.Count
    ncols = INVOICE_COL
    for row in range(INVOICE_HEADER_ROW, nrows):
        i = row - INVOICE_HEADER_ROW
        j = 0
        for col in range(1, ncols):
            value = ws.Cells.Item(row, col).Value
            if j == 0:
                if value == None:
                    break
                else:
                    invoice_data[i][j] = value
            else:
                invoice_data[i].append(value)
            j += 1
        else:
            invoice_data.append([0])
            continue
        break
    invoice_data.pop(-1)
    return(invoice_data)

# 請求書データ存在チェック
def is_invoice_exist(invoice_data, id):
    for row in invoice_data:
        if row[1] == id:
            return(True)
    return(False)

# 配列内の請求書データを取得
def get_invoice_from_list(invoice_list, id):
    for iv in invoice_list:
        if iv.id == id:
            return(iv)
    return(None)

# 請求書PDF生成
def create_invoice_pdf(app, wb, ws, invoice, pdf_output_path="../pdf"):
    # 請求書テンプレートコピー
    app.DisplayAlerts = False
    ws.Copy(None, wb.Sheets(wb.Sheets.Count))
    iv_doc = wb.Sheets(wb.Sheets.Count)
    iv_doc.Name = 'invoice_work'

    # 請求書作成
    iv_doc.Cells(COMPANY_NAME_ROW, COMPANY_NAME_COL).Value = invoice.company_name
    iv_doc.Cells(PERSONNEL_ROW, PERSONNEL_COL).Value = invoice.personnel
    iv_doc.Cells(INVOICE_ID_ROW, INVOICE_ID_COL).Value = invoice.id
    today_date = datetime.date.today()
    iv_doc.Cells(TODAY_ROW, TODAY_COL).Value = pdate_to_str(today_date)
    iv_doc.Cells(TITLE_ROW, TITLE_COL).Value = invoice.title
    iv_doc.Cells(DUE_DATE_ROW, DUE_DATE_COL).Value = invoice.due_date
    detail_row = DETAIL_START_ROW
    i = 0
    for iv_item in invoice.invoice_details:
        iv_doc.Cells(detail_row + i, DETAIL_NUM_COL).Value = i + 1
        iv_doc.Cells(detail_row + i, PRODUCT_NAME_COL).Value = iv_item.product_name
        iv_doc.Cells(detail_row + i, UNIT_NUM_COL).Value = iv_item.unit_num
        iv_doc.Cells(detail_row + i, UNIT_PRICE_COL).Value = iv_item.unit_price
        i += 1
    iv_doc.Cells(DESCRIPTION_ROW, DESCRIPTION_COL).Value = invoice.description
    # 請求書出力
    invoice_path = str(pathlib.Path(pdf_output_path + '/' + invoice.company_name + pdate_to_str(today_date) + '.pdf').resolve())
    iv_doc.ExportAsFixedFormat(0,invoice_path) 
    # 請求書シート削除
    wb.Worksheets(wb.Sheets.Count).Delete()

    wb.Save()
    app.DisplayAlerts = True