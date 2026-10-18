# 業務自動化のデモ集

「手作業の転記・集計・通知」を自動化するツールのデモです。
GAS（Google Apps Script）を軸に、生成AI API を組み合わせています。

## デモ一覧

| テーマ | 技術 | フォルダ |
|---|---|---|
| 領収書の写真をAIで読み取り、経費表に自動で記録 | GAS、Gemini API | [receipt-ocr](demos/receipt-ocr/) |
| フォームの問い合わせを一覧に転記して、担当者にメールで通知 | GAS、Googleフォーム、Gmail | [form-notify](demos/form-notify/) |
| アンケートの回答を毎週集計してメールで送り、推移をグラフで見せる | GAS、Googleフォーム、Gmail、スプレッドシートのグラフ | [form-report](demos/form-report/) |
| 請求書をインボイス対応のPDFとGmailの下書きにし、入金の照らし合わせまで自動に | GAS、Googleドライブ、Gmail | [invoice](demos/invoice/) |

## フォルダ構成

- `demos/<名前>/` … 各デモ。中に `README.md`（Before/After・使い方）とコード。GASのデモは [clasp](https://github.com/google/clasp) で管理しているので、`appsscript.json` も一緒に置いています
- `lib/gas/` … GASの共通部品（2つ以上のデモで使う処理）

## デモのデータについて

- デモに出てくる領収書・お客様情報などは、すべて架空のものです
- AIを使うためのパスワード（APIキー）は、プログラムの中に書かず、外から見えない場所に保管しています
