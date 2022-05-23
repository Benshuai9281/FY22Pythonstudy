D3 = "Konica"
D5 = "Minolta"
D3_5 = "KonicaMinolta"

class Konicaminolta(object):
    def __init__(self) -> None:
        pass

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

if __name__ == "__main__":
    newkm = Konicaminolta()
    newkm.kmcalc()