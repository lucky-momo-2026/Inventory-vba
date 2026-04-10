from numpy import size
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

    plt.figure(figsize=(10, 5), dpi=120)  #横に二つ並べるため横長にする

    plt.subplot(1, 2, 1)  #１行２列の１つめ（左）

     #棒グラフを作成して棒の情報を受け取る
    bars = plt.bar(labels, values, color=["red", "orange"])  #色指定

    plt.ylim(0, max(values) + 1.5)  #上に余白を作る
    plt.yticks(range(0, max(values) +1))

    plt.xticks(fontsize=14)  #横軸ラベルの文字を大きくする

    #棒の上に件数を表示する
    for bar in bars:
        height = bar.get_height()  #棒の高さ（件数）を取得
        plt.text(
            bar.get_x() + bar.get_width() / 2,  #棒の中央値
            height + 0.05,  #棒の高さの位置
            f"{int(height)}",  #表示する文字
            ha="center",  #横方向は中央ぞろえ
            va="bottom",  #縦方向は棒の上に置く
            fontsize=14
        )

    #タイトルと軸ラベル
    plt.title("在庫状況の件数集計", fontsize=16)
    plt.ylabel("件", rotation=0, labelpad=20, fontsize=14) #ratation横向き/labelpad左に離す

    #円グラフ用の件数データ
    sizes = [out_of_stock_count, low_stock_count]

    plt.subplot(1, 2, 2)  #1行2列の2つ目（右）

    #円グラフを書く
    plt.pie(
        sizes,
        labels = ["在庫切れ", "在庫少"],  
        autopct = "%1.1f%%",  #割合を表示
        startangle = 90,  #上から始める
        colors = ["red", "orange"]  #在庫切れ赤　在庫少オレンジ
    )

    plt.title("在庫の状況を割合", fontsize=16)  #タイトル
    plt.axis("equal")  #円をきれいな丸にする

    plt.suptitle("在庫状況レポート", fontsize=20)

    plt.tight_layout(rect=[0, 0, 1, 0.93])  #２つのグラフが重ならないように整える
    plt.savefig("inventory_summary.png", dpi=300)  #合体グラフを一枚で保存

    plt.show()  #画面に表示

    # メモリ節約のためグラフを閉じる
    plt.close()

#件数集計グラフを画像として保存する
save_stock_graph(out_of_stock_count, low_stock_count)
    
print("グラフを inventory_graph.png に保存しました")    
