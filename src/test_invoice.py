import unittest
import datetime
import os
import win32com.client
from pathlib import Path
import glob
from invoice import Invoice, InvoiceDetail
import invoice_util

class TestInvoice(unittest.TestCase):
    def setUp(self):
        self.app = win32com.client.Dispatch("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        abs_path = str(Path(r"../invoice.xlsm").resolve())
        self.wb = self.app.Workbooks.Open(abs_path)
        self.ws = self.wb.WorkSheets("請求データ")

    def test_get_excel_data(self):
        iv_data = invoice_util.get_excel_data(self.ws)
        self.assertEqual(iv_data[0][3],"XXX株式会社")
    
    def test_is_invoice_exist(self):
        iv_data = invoice_util.get_excel_data(self.ws)
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

        self.assertEqual(invoice_util.is_invoice_exist(invoices, 1), True)
        self.assertEqual(invoice_util.is_invoice_exist(invoices, 5), False)

    def test_add_detail(self):
        iv_data = invoice_util.get_excel_data(self.ws)
        iv = Invoice(iv_data[0])
        iv_detail = InvoiceDetail(iv_data[0])
        iv.add_detail(iv_detail)
        self.assertEqual(iv.invoice_details[0].product_name, "HDD")

    def test_get_invoice_from_list(self):
        iv_data = invoice_util.get_excel_data(self.ws)
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

        iv1 = invoice_util.get_invoice_from_list(invoices, 2)
        self.assertEqual(iv1.title, "PC USBケーブル")
        iv2 = invoice_util.get_invoice_from_list(invoices, 5)
        self.assertEqual(iv2, None) 

    def test_excel_date_to_str(self):
        self.assertEqual(invoice_util.excel_date_to_str(43951),"2020年4月30日")

    def test_pdate_to_str(self):
        sdate = datetime.datetime(2020, 3, 11)
        self.assertEqual(invoice_util.pdate_to_str(sdate),"2020年3月11日")        

    def test_create_pdf(self):
        invoice_doc = self.wb.WorkSheets("請求書テンプレート")
        iv_data = invoice_util.get_excel_data(self.ws)
        iv = Invoice(iv_data[0])
        iv_detail = InvoiceDetail(iv_data[0])
        iv.add_detail(iv_detail)
        invoice_util.create_invoice_pdf(self.app, self.wb, invoice_doc, iv, "../pdf")
        f_list = glob.glob('../pdf/XXX株式会社*.pdf')
        self.assertEqual(len(f_list), 1)

    def tearDown(self):
        r_list = glob.glob('../pdf/*')
        for r_file in r_list:
            os.remove(r_file)
        self.wb.Close()
        self.app.Quit()

if __name__ == "__main__":
    unittest.main()