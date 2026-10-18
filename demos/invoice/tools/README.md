# tools

開発用の道具。お客様には渡さない（Apps Script には送らない）。

- `make_buttons.py` … 「操作」シートのボタンの画像を作り、`Button.gs` に書き出す。`python tools/make_buttons.py`
- `DemoData.gs` … デモ用の架空のデータ（設定・取引先・毎月の請求・明細）を入れる `loadDemoData`。ツールのシートと記録を作り直し、前のテストのPDF・CSVはゴミ箱へ移す。使うときだけデモのフォルダに戻して `clasp push` し、エディタから実行する。終わったら `tools/` に戻す
- `DemoCsv.gs` … 請求書を作ったあとに、一覧の金額から架空のネットバンクのCSV（Shift_JIS）を作る `makeDemoBankCsv`。使い方は `DemoData.gs` と同じ

会社・取引先・口座はすべて架空。登録番号は、検査数字が合わない番号（実在しえない）にしている。
