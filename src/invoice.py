# invoice.py

# 請求書明細データクラス
class InvoiceDetail:
    def __init__(self, data):
        # 商品名
        self.product_name = data[6]
        # 単価
        self.unit_price = data[7]
        # ユニット数
        self.unit_num = int(data[8])

# 請求書データクラス
class Invoice:
    def __init__(self, data):
        # 請求書ID
        self.id = int(data[1])
        # 会社ID
        self.company_id = int(data[2])
        # 会社名
        self.company_name = data[3]
        # 担当者名
        self.personnel = data[4]
        # 請求書タイトル
        self.title = data[5]
        # 発行日
        self.date = data[10]
        # 支払い期限
        self.due_date = data[12]
        # 備考 
        self.description = data[13]
        # 請求書明細
        self.invoice_details = []
    
    # 請求書データに請求書明細データを追加
    def add_detail(self, invoice_detail):
        self.invoice_details.append(invoice_detail)