#該当数値で割り切れる場合の表示文字を定義する。
D3 = "Konica"
D5 = "Minolta"
D3_5 = "KonicaMinolta"

#Konicaminoltaのクラスを定義。
class Konicaminolta(object):
    def __init__(self) -> None:
        pass
#数値で割り切れるかを判断して、文字を表示するメソッドを定義。
    def kmcalc(self) -> None:
        for i in range(1,101):
            if i % 5 == 0 and i % 3 == 0:
                print (f"{i}:{D3_5}")
            elif i % 5 == 0:
                print (f"{i}:{D5}")
            elif i % 3 == 0:
                print (f"{i}:{D3}")
            else:
                print (f"{i}:")
#このファイルからプログラムを起動する場合は、Konicaminoltaクラスのインスタンスを作成し、メソッドを実行する。
if __name__ == "__main__":
    newkm = Konicaminolta()
    newkm.kmcalc()