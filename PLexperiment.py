from monochromatorCM110 import CM110
import string
import serial
import time

if __name__ == "__main__":
    CM = CM110()
    high = input("input high wavelength:")
    low = input("input low wavelength")
    thisbyte = low
    while thisbyte == high:
        CM.gotoWave(thisbyte)
        time.sleep(0.2)
        thisbyte = thisbyte + 1
    