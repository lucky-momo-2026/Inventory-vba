import pandas as pd

#cvsを読み込む（文字化け対策でencoding指定）
df = pd.read_csv('result.csv', encoding='cp932')

#データの確認
print('=== データ確認 ===')
print(df)

#判定ごとの研修を集計する
out_of_stock_count =(df['判定'] == '在庫切れ').sum()
low_stock_count = (df['判定'] == '在庫少').sum()

print('\n=== 件数集計 ===')
print(f'在庫切れ件数：{out_of_stock_count}件')
print(f'在庫少件数：{low_stock_count}件')

