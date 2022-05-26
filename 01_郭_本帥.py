#正規表現のモジュールを導入する。
import re

#表示文字を定義する。
WELCOMEINFORMATION = """Welcome to the leap year check tool!
Please input years you want to check (split with ","):
Example: 1000,1992,2009,2001
==>"""

#Yearのクラスを定義。
class Year(object):
    def __init__(self,info) -> None:
        self.info = info
#入力、及び入力された文字列のフォーマットが正しいかを判断するメソッドを定義。
    def yearinput(self) -> list:
        while True:
            _yearstr = input(self.info)
            #入力した内容は、一桁以上の数値（d）からスタートし、いくつの「,」と数値の組合を付いている文字列をマッチングして、正しい入力を判断する。
            if re.match(r'^\d+(,\d+)*$',_yearstr):
                #「,」で文字列を分割する。
                _yearlist=_yearstr.split(",")
                return _yearlist
            else:
                print("The input data format is incorrect, please input again.\n")
                
#数値で割り切れるかを判断して、文字を表示するメソッドを定義。
    def yearcheck(self,year) -> None:
        yearlistforcheck = list(map(int,year))
        for y in yearlistforcheck:
            if y % 400 == 0:
                print (f"{y} is a leap year")
            elif y % 100 == 0:
                print (f"{y} is not a leap year")
            elif y % 4 == 0:
                print (f"{y} is a leap year")
            else:
                print (f"{y} is not a leap year")

#このファイルからプログラムを起動する場合は、Yearクラスのインスタンスを作成し、メソッドを実行する。
if __name__ == "__main__":
    checkyear = Year(WELCOMEINFORMATION)
    checkyear.yearcheck(checkyear.yearinput())