import pandas as pd

# DataFrameの作成

df = pd.DataFrame({

    '出身': ['東京','神奈川','福岡','北海道','愛知'],

    '年代': [20,20,30,40,50],

    '継続年数': [2,3,4,10,21],

    '在籍': ['八王子','瑞穂','八王子','高槻','丸の内']},

    )

for col_name, col in df.iteritems():
  for row_name in col.index:
    if df.loc[row_name, col_name] == "八王子":
      df.loc[row_name, col_name] = "梅田"

for row_name, row in df.iterrows():
  for col_name in row.index:
    if df.loc[row_name, col_name] == "八王子":
      df.loc[row_name, col_name] = "梅田"

print(df)
