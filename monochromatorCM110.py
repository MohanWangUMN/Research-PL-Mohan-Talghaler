"""
-------------------------------------------------------------------------
This a class for the CM110 monochromator from Spectral Products
(the little white box.) Communication is only done by serial.

 MODIFICATION HISTORY

 Maintainer: Merlin Mah
 [ 2/ 2/2023] More Python3 fixes.
 [12/ 4/2022] Very quick and dirty patches to run under Python3.
              Renamed class from DK110 to CM110.

 Author: Phil Armstrong
 [ 6/13/2012] The class only has the capability to change the wavelength.
  In the "gotoWave" there is code for converting from decimal to hexadecimal
  low byte and high byte which is how the CM110 communicates. Later this
  should be added to a seperate function which can be called by the other
  functions. The units are in angstroms.
--------------------------------------------------------------------------
"""

"""
edit by Mohan Wang for PL experiment
2/14/2023
version 0.0. 
original py file copy from smb://files.umn.edu/ece/Research/Talghaler/JTO/python

"""

import serial
import string
import gizmos
import time


class CM110():

        def __init__(self, comPort=2):
            self.commport = comPort
            # self.ser = serial.Serial(self.commport, timeout=1)
            self.ser = serial.Serial(self.commport, baudrate=9600, bytesize=8, parity='N', stopbits=1, xonxoff=0, rtscts=1, timeout=1, write_timeout=1)
            self.giz = gizmos.Gizmos()


        # Set the wavelength in angstroms
        def gotoWave(self, wave):
            lo = wave & 0x00ff
            hi = (wave & 0xff00) >> 8
            for thisbyte in [0x10, hi, lo]:
                self.ser.write(chr(thisbyte).encode(encoding="latin-1"))
            time.sleep(0.5)
            first = self.ser.read(4).decode(encoding="latin-1")
            for thisbyte in [56, 00]:
                self.ser.write(chr(thisbyte).encode(encoding="latin-1"))
            # self.ser.write(chr(56) + chr(00))
            second = self.ser.read(2)
            third = self.ser.read(2)
            convHex = second[1] + (second[0] << 8)
            # print(f"Wavelength is at {int(convHex)/10} + nm.") # diagnostics
            return int(convHex) # Still in Angstroms


        def reset(self):
            for thisbyte in [0xff, 0xff, 0xff]:
                self.ser.write(chr(thisbyte).encode(encoding="latin-1"))
            # No return message noted.
            print('\nMonochromator resetting.\n')



        #close serial port when done
        def closeport(self):
            self.ser.close()
