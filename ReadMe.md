## 請求書作成自動化Pythonスクリプト

Excelファイルにある請求書データからPDFフォーマットの請求書を生成する。  

## 実行環境
ExcelがインストールされているWindows OS PC  
Python3 Windows版  
PowerShellまたはコマンドプロンプトを管理者権限で立ち上げて実行する

## ファイル  
- invoice.xlsm  
  請求書データ入力シートと請求書テンプレート
- requirements.txt  
  Pythonモジュール設定ファイル
- src/app.py  
  メインプログラム
- src/invoice.py  
  Invoiceクラスファイル  
- src/invoice_util.py  
  ユーティリティ関数ファイル
- src/test_invoice.py  
  テストファイル
- pdf  
  請求書PDFファイル生成ディレクトリ

## テスト実行方法
python -m unittest test_invoice.py