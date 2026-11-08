# KAKERU わたしのいまここ

アスリート向けセルフコンディションチェック（アスユメLabo.）。

| パス | 内容 |
| --- | --- |
| `index.html` | 受検画面（GitHub Pagesで公開） |
| `gas/code.gs` | Apps Script（スプレッドシート「KAKERU受検結果」に紐づくスクリプトの コード.gs） |
| `gas/Index.html` | Apps Script側の受検画面（予備） |

- 回答は `index.html` から Apps Script のウェブアプリ（doPost）に送られ、スプレッドシートに記録、本人へPDFレポートをメール送付します。
- `gas/` を変更したら、Apps Scriptに貼り付けて［デプロイを管理］→ 新バージョンで公開します（URLは変えない）。
