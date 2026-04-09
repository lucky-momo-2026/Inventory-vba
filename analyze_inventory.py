import pandas as pd
import matplotlib.pyplot as plt #グラフを作る

#cvsを読み込む（文字化け対策でencoding指定）
df = pd.read_csv("result.csv", encoding="cp932")

#データの確認
print("=== データ確認 ===")
print(df)

#判定ごとの研修を集計する
out_of_stock_count =(df["判定"] == "在庫切れ").sum()
low_stock_count = (df["判定"] == "在庫少").sum()

print("\n=== 件数集計 ===")
print(f"在庫切れ件数：{out_of_stock_count}件")
print(f"在庫少件数：{low_stock_count}件")

def save_stock_graph(out_of_stock_count, low_stock_count):
    """在庫切れ件数と在庫少件数を棒グラフにして保存する"""

    #日本語が文字化けしにくいようにフォントを設定
    plt.rcParams["font.family"] = "Yu Gothic"

    labels = ["在庫切れ", "在庫少"]
    values = [out_of_stock_count, low_stock_count]

    #棒グラフを作成
    plt.figure(figsize=(6, 4))
 
    bars = plt.bar(labels,values)  #棒グラフを作成して棒の情報を受け取る

    #棒の上に件数を表示する
    for bar in bars:
        height = bar.get_height()  #棒の高さ（件数）を取得
        plt.text(
            bar.get_x() + bar.get_width() / 2,  #棒の中央値
            height,  #棒の高さの位置
            f"{int(height)}",  #表示する文字
            ha="center",  #横方向は中央ぞろえ
            va="bottom"  #縦方向は棒の上に置く
        )

    #タイトルと軸ラベル
    plt.title("在庫状況の件数集計")
    plt.ylabel("件数", rotation=0, labelpad=20) #ratation横向き/labelpad左に離す

    # 画像として保存
    plt.savefig("inventory_graph.png")

    # メモリ節約のためグラフを閉じる
    plt.close()

#件数集計グラフを画像として保存する
save_stock_graph(out_of_stock_count, low_stock_count)
    
print("グラフを inventory_graph.png に保存しました")    
