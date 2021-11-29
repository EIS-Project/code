"""
   DWF Python Example
   Author:  Digilent, Inc.
   Revision:  2019-07-12

   Requires:                       
       Python 2.7, 3
"""

from ctypes import *
from dwfconstants import *
import math
import time
import sys
import numpy
import matplotlib.pyplot as plt

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

# this option will enable dynamic adjustment of analog out settings like: frequency, amplitude...
dwf.FDwfDeviceAutoConfigureSet(hdwf, c_int(3)) 

sts = c_byte()
steps = 151
start = 1e2
stop = 1e5
reference = 1e3
amplitude = 1

t1 = c_double()
t2 = c_double()

print("Frequency: "+str(start)+" Hz ... "+str(stop/1e3)+" kHz Steps: "+str(steps))
dwf.FDwfAnalogImpedanceReset(hdwf)
dwf.FDwfAnalogImpedanceModeSet(hdwf, c_int(0)) # 0 = W1-C1-DUT-C2-R-GND, 1 = W1-C1-R-C2-DUT-GND, 8 = AD IA adapter
dwf.FDwfAnalogImpedanceReferenceSet(hdwf, c_double(reference)) # reference resistor value in Ohms
dwf.FDwfAnalogImpedanceFrequencySet(hdwf, c_double(start)) # frequency in Hertz
dwf.FDwfAnalogImpedanceAmplitudeSet(hdwf, c_double(amplitude)) # 1V amplitude = 2V peak2peak signal
dwf.FDwfAnalogImpedanceConfigure(hdwf, c_int(1)) # start
time.sleep(2)

rgHz = [0.0]*steps
rgGaC2 = [0.0]*steps
rgPhC2 = [0.0]*steps
rgCs = [0.0]*steps
for i in range(steps):
    hz = stop * pow(10.0, 1.0*(1.0*i/(steps-1)-1)*math.log10(stop/start)) # exponential frequency steps
    rgHz[i] = hz
    dwf.FDwfAnalogImpedanceFrequencySet(hdwf, c_double(hz)) # frequency in Hertz
    time.sleep(0.01) 
    dwf.FDwfAnalogImpedanceStatus(hdwf, None) # ignore last capture since we changed the frequency
    while True:
        if dwf.FDwfAnalogImpedanceStatus(hdwf, byref(sts)) == 0:
            dwf.FDwfGetLastErrorMsg(szerr)
            print(str(szerr.value))
            quit()
        if sts.value == 2:
            break
    gain2 = c_double()
    phase2 = c_double()
    dwf.FDwfAnalogImpedanceStatusInput(hdwf, c_int(1), byref(gain2), byref(phase2)) # relative to Channel 1, C1/C#
    rgGaC2[i] = 20.0*math.log10(abs(gain2.value-1.0))
    rgPhC2[i] = phase2.value*180/math.pi
    cs = c_double()
    dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceSeriesCapactance, byref(cs)) 
    rgCs[i] = cs.value
    

dwf.FDwfAnalogImpedanceConfigure(hdwf, c_int(0)) # stop
dwf.FDwfDeviceClose(hdwf)


plt.subplot(311)
plt.plot(rgHz, rgGaC2, color='blue')
ax = plt.gca()
ax.set_title("{0:.2f}Hz : {1:.2f}dB {2:.2f}* {3:.2f}nF".format(rgHz[0], rgGaC2[0], rgPhC2[0], rgCs[0]*1e9))
ax.set_xscale('log')
plt.subplot(312)
plt.plot(rgHz, rgPhC2)
ax = plt.gca()
ax.set_xscale('log')
plt.subplot(313)
plt.plot(rgHz, rgCs)
ax = plt.gca()
ax.set_xscale('log')
plt.show()

