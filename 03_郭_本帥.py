from numpy import object_
import pandas as pd

col_author = ["詠み人", "天皇", "姫", "坊主"]
index_author = [
["天智天皇", "●", "", ""],
["持統天皇", "●", "●", ""],
["小野小町", "", "●", ""],
["喜撰法師", "", "", "●"],
["蝉丸", "", "", "●"],
["柿本人麻呂", "", "", ""],
["山部赤人", "", "", ""],
["猿丸太夫", "", "", ""],
["大伴家持", "", "", ""],
["安倍仲麻呂", "", "", ""]]

df_author = pd.DataFrame(index_author, columns=col_author)


col_waka = ["上の句", "下の句", "詠み人"]
index_waka = [
["秋の田のかりほの庵の苫を荒み", "わがころも手は露に濡れつつ", "天智天皇"],
["春すぎて夏来にけらし白たへの", "ころもほすてふあまの香具山", "持統天皇"],
["あしひきの山鳥の尾のしだり尾の", "ながながし夜をひとりかも寝む", "柿本人麻呂"],
["田子の浦にうちいでて見れば白たへの", "富士の高嶺に雪は降りつつ", "山部赤人"],
["奥山にもみぢ踏み分け鳴く鹿の", "声聞く時ぞ秋は悲しき", "猿丸太夫"],
["かささぎの渡せる橋に置く霜の", "白きを見れば夜ぞふけにける", "大伴家持"],
["あまの原ふりさけ見ればかすがなる", "み笠の山にいでし月かも", "安倍仲麻呂"],
["わが庵は都のたつみしかぞ住む", "世を宇治山と人は言ふなり", "喜撰法師"],
["花の色はうつりにけりないたづらに", "わが身世にふるながめせしまに", "小野小町"],
["これやこの行くも帰るも別れては", "知るも知らぬも逢坂の関", "蝉丸"]]

df_waka = pd.DataFrame(index_waka, columns=col_waka)


df_hyakunin_isshu = pd.merge(df_waka, df_author, on="詠み人", how="inner")

df_hyakunin_isshu["歌"] = df_hyakunin_isshu["上の句"].values + df_hyakunin_isshu["下の句"].values
del df_hyakunin_isshu["上の句"]
del df_hyakunin_isshu["下の句"]

df_hyakunin_isshu["属性"] = ""

for i in df_hyakunin_isshu.index:
  if df_hyakunin_isshu["天皇"][i] == "●":
    df_hyakunin_isshu["属性"][i] = df_hyakunin_isshu["属性"][i] + ";天皇"
  if df_hyakunin_isshu["姫"][i] == "●":
    df_hyakunin_isshu["属性"][i] = df_hyakunin_isshu["属性"][i] + ";姫"
  if df_hyakunin_isshu["坊主"][i] == "●":
    df_hyakunin_isshu["属性"][i] = df_hyakunin_isshu["属性"][i] + ";坊主"
  if df_hyakunin_isshu["属性"][i] != "":
      df_hyakunin_isshu["属性"][i] = df_hyakunin_isshu["属性"][i][1:]
del df_hyakunin_isshu["天皇"]
del df_hyakunin_isshu["姫"]
del df_hyakunin_isshu["坊主"]

#print (df_hyakunin_isshu)
df_hyakunin_isshu_bose = df_hyakunin_isshu.loc[[True if i.find("坊主") != -1 else False for i in df_hyakunin_isshu['属性'].values]]
#print (df_hyakunin_isshu_bose)
df_hyakunin_isshu_bose.to_excel('./bose.xlsx', index=False, sheet_name='百人一首（坊主）')