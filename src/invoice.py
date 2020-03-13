# invoice.py

class InvoiceDetail:
    def __init__(self, data):
        self.product_name = data[6]
        self.unit_price = data[7]
        self.unit_num = int(data[8])

class Invoice:
    def __init__(self, data):
        self.id = int(data[1])
        self.company_id = int(data[2])
        self.company_name = data[3]
        self.personnel = data[4]
        self.title = data[5]
        self.date = data[10]
        self.due_date = data[12]
        self.description = data[13]
        self.invoice_details = []

    def add_detail(self, invoice_detail):
        self.invoice_details.append(invoice_detail)