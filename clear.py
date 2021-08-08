from os import system,name
import time
def clear():
    if name=="nt":
        _=system("cls")
    else:
        _=system("clear")
time.sleep(0.5)
clear()