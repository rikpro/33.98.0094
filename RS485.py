"""
This module defines the send functionality to send commands to LED displays.

@author: Stijn Deroo
@since: 7-mrt-2013
@copyright: Televic Rail

"""

import socket
import threading
import serial
import Queue
import time

class RS485():
    def __init__(self, port, baudrate=115200):
        self.ser = serial.Serial(port, baudrate, timeout=1)
        self.ser.bytesize = serial.EIGHTBITS
        self.ser.parity = serial.PARITY_NONE
        self.ser.stopbits = serial.STOPBITS_ONE
        self.ser.xonxoff  = False
        self.ser.rtscts   = False
        self.ser.writeTimeout = 0.1
        self.ser.interCharTimeout = 0.1
        
    def openPort(self):
        self.ser.open()
        
    def closePort(self):
        self.ser.close()

    def pingPong(self, dataOut = "Televic"):
        self.tx(dataOut)
        time.sleep(0.2)      
        dataIn = self.rx()
        return dataIn == dataOut       
        
    def tx(self, dataOut = "Televic"):
#         self.ser.flushInput()       #reset_input_buffer()
#         self.ser.flushOutput()      #reset_output_buffer()
        self.ser.write("".join(map(chr, dataOut)))  
        self.ser.flushInput()       #reset_input_buffer()
        self.ser.flushOutput()      #reset_output_buffer()
        return      
    
    def rx(self, numberOfBytes = None): 
        if numberOfBytes == None:
            dataIn = self.ser.read(self.ser.inWaiting())
        else:
            dataIn = self.ser.read(numberOfBytes)
        return bytearray(dataIn)
    
    def emptyBuffer(self):
        self.ser.flushInput()       #reset_input_buffer()
        self.ser.flushOutput()      #reset_output_buffer()
        
    def getSendData(self):
        dataOut = bytearray()
        dataOut.append(0x02)
        dataOut.append(0x20)
        dataOut.append(0xab)
        dataOut.append(0xcd)
        dataOut.append(0x08)
        dataOut.append(0x00)
        dataOut.append(0x5f)
        dataOut.append(0x06)
        
        return dataOut
    
    def getReceivedData(self):
        dataIn = bytearray()
        dataIn.append(0x02)
        dataIn.append(0x20)
        dataIn.append(0x31)
        dataIn.append(0xab)
        dataIn.append(0xcd)
        dataIn.append(0x08)
        dataIn.append(0x00)
        dataIn.append(0x21)
        dataIn.append(0x7d)
        dataIn.append(0x23)
        dataIn.append(0x0a)
        dataIn.append(0x03)
        
        return dataIn