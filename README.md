\# 📊 在庫状況グラフ作成スクリプト



VBAで抽出した在庫データ（`result.csv`）を読み込み、在庫切れ・在庫少の件数を集計して棒グラフ・円グラフで可視化するPythonスクリプトです。



\---



\## 📋 機能一覧



\- 在庫切れ・在庫少の件数をターミナルに表示

\- 棒グラフで件数を可視化（棒の上に件数を自動表示）

\- 円グラフで在庫状況の割合を可視化

\- 棒グラフ・円グラフを1枚にまとめた画像を `inventory\_summary.png` に保存



\---



\## 🗂️ ファイル構成



```

inventory-check/

├── analyze\_inventory.py     # このスクリプト

├── result.csv               # 入力データ（VBAマクロが出力したもの）

└── inventory\_summary.png    # 出力：在庫状況グラフ画像

```



\---



\## 📦 必要なライブラリ



```bash

pip install pandas matplotlib

```



\---



\## 📄 入力CSVの形式



```csv

商品名,在庫数,判定

ノート,0,在庫切れ

ペン,3,在庫少

消しゴム,0,在庫切れ

ファイル,2,在庫少

```



> `result.csv` はVBAマクロを先に実行すると自動生成されます。



\---



\## 🚀 使い方



```bash

python analyze\_inventory.py

```



\### 出力例（ターミナル）



```

=== データ確認 ===

（result.csv の内容が表示されます）



=== 件数集計 ===

在庫切れ件数：2件

在庫少件数：2件

```

> ※上記はサンプルデータによる出力例です



\### 出力画像（inventory\_summary.png）



棒グラフと円グラフを1枚にまとめたレポート画像が生成されます。

!\[在庫状況グラフ](inventory\_summary.png)



\---



\## 🛠️ 開発環境



\- Python 3.x

\- pandas

\- matplotlib



\---



\## 👤 作者



\[lucky-momo-2026](https://github.com/lucky-momo-2026)

