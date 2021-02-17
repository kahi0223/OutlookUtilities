# latest csv mailer

## overview

Windows用のOutlook専用の自動CSVファイル更新・配信メール作成プログラム。

最新のCSVファイルがメールで送られてくる場合に、ローカルに保存しているCSVファイルにmergeして、メールを配信する。

定期的に送られてくるRawCSVデータをローカルのCSVファイルにmergeして、複数人に配信することを想定。

`config.py`で基本設定ができる。
`main.py`で各関数の引数を変更すれば送信先なども変更できる。

```shell
.
│  clock.py # 時刻周りを扱うクラス。タイムゾーンが使用できるようになっている。
│  config.py # 設定ファイル
│  csv_hander.py # CSVを扱うクラス
│  emailer.py # メール検索・作成クラス
│  excel_hander.py # (Excelを配信するためのクラス。マクロを実行できたが、できなくなったので未使用)
└─ main.py # 実行ファイル
```

## 実行方法 

```shell
python main.py '検索したいメールのタイトル名'
```

## output例

```shell
python3 main.py Subject
```

```shell
#outlook

TO: aaa@XXX.com; bbb@XXX.com; aaa@XXX.com; bbb@XXX.com
CC: AAA@XXX.com; BBB@XXX.com; AAA@XXX.com; BBB@XXX.com

SUBJECT: LatestMerged_Subject_2021/02/15(UTC+09:00)
ATTACHMENT: MAIN.csv

BODY:
Dear XXX,

XXXXX.
SAVED folder 
(C:\Users\FOLDERNAME\CSV\MAIN.csv)
```