from monochromatorCM110 import CM110
import string
import serial
import time

if __name__ == "__main__":
    CM = CM110()
    if CM.echo() is not True:
        print("fail to connect with CM110.\n")
        exit()
    CM.reset()
    while 1:
        print('1. calibration\n2. setwavelength\n3. reset CM110\n4. check connection from CM110\n')
        i = input()
        if i is 1:
            wavelength = input("input wavelength for calibration")
            CM.calibrate()
        elif i is 2:
            wavelength = input("set ur wavelength")
            CM.gotoWave(wavelength)
        elif i is 3:
            CM.reset()
        elif i is 4:
            if CM.echo(): 
                print("\n")
            else:
                print("fail to connect with CM110!\n")

        if i == "exit":
            exit()


        