import re

WELCOMEINFORMATION = """Welcome to the leap year check tool!
Please input years you want to check (split with ","):
Example: 1000,1992,2009,2001
==>"""

class Year(object):
    def __init__(self,info) -> None:
        self.info = info

    def yearinput(self) -> list:
        while True:
            _yearstr = input(self.info)
            if re.match(r'^\d+(,\d+)*$',_yearstr):
                _yearlist=_yearstr.split(",")
                return _yearlist
            else:
                print("The input data format is incorrect, please input again.\n")

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

if __name__ == "__main__":
    checkyear = Year(WELCOMEINFORMATION)
    checkyear.yearcheck(checkyear.yearinput())