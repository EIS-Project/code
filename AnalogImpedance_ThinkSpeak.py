"""
   DWF Python Example
   Author:  Digilent, Inc.
   Revision:  2019-10-21

   Requires:                       
       Python 2.7, 3
   Desciption:
   Performs impedance measurements and pushes to ThingSpeak.com
"""

from ctypes import *
from dwfconstants import *
import math
import time
import sys
import numpy
import requests 

url = "https://api.thingspeak.com/update?api_key=8C############BU"

if sys.platform.startswith("win"):
    dwf = cdll.LoadLibrary("dwf.dll")
elif sys.platform.startswith("darwin"):
    dwf = cdll.LoadLibrary("/Library/Frameworks/dwf.framework/dwf")
else:
    dwf = cdll.LoadLibrary("libdwf.so")

version = create_string_buffer(16)
dwf.FDwfGetVersion(version)
print("DWF Version: "+str(version.value))

hdwf = c_int()
szerr = create_string_buffer(512)
print("Opening first device")
dwf.FDwfDeviceOpen(c_int(-1), byref(hdwf))

if hdwf.value == hdwfNone.value:
    dwf.FDwfGetLastErrorMsg(szerr)
    print(str(szerr.value))
    print("failed to open device")
    quit()

sts = c_byte()
frequnecy = 1e3
reference = 1e6
capacitance = c_double()
resistance = c_double()
reactance = c_double()

print("Reference: "+str(reference)+" Ohm  Frequency: "+str(frequnecy/1e3)+" kHz")
dwf.FDwfAnalogImpedanceReset(hdwf)
dwf.FDwfAnalogImpedanceModeSet(hdwf, c_int(8)) # 0 = W1-C1-DUT-C2-R-GND, 1 = W1-C1-R-C2-DUT-GND, 8 = AD IA adapter
dwf.FDwfAnalogImpedanceReferenceSet(hdwf, c_double(reference)) # reference resistor value in Ohms
dwf.FDwfAnalogImpedanceFrequencySet(hdwf, c_double(frequnecy)) # frequency in Hertz
dwf.FDwfAnalogImpedanceAmplitudeSet(hdwf, c_double(1))
dwf.FDwfAnalogImpedanceConfigure(hdwf, c_int(1)) # start
time.sleep(1)

dwf.FDwfAnalogImpedanceStatus(hdwf, None) # ignore last capture, force a new one

print("Press Ctrl+C to stop...")

try:
    for i in range(1000) :
        time.sleep(60) # seconds
        if dwf.FDwfAnalogImpedanceStatus(hdwf, byref(sts)) == 0:
            dwf.FDwfGetLastErrorMsg(szerr)
            print(str(szerr.value))
            quit()
        if sts.value != 2:
            print("Measurement not done")
            continue
        dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceResistance, byref(resistance))
        dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceReactance, byref(reactance))
        dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceSeriesCapactance, byref(capacitance))
        print(str(i)+" Resistance: "+str(resistance.value)+" Ohm  Reactance: "+str(reactance.value/1e6)+" MOhm  Capacitance: "+str(capacitance.value*1e12)+" pF")
        r = requests.get(url+"&field1="+str(resistance.value)+"&field2="+str(reactance.value/1e6)+"&field3="+str(capacitance.value*1e12))
        if r.status_code != 200:
            print(r)
            break

except KeyboardInterrupt: # Ctrl+C
    pass

dwf.FDwfAnalogImpedanceConfigure(hdwf, c_int(0)) # stop
dwf.FDwfDeviceClose(hdwf)
