"""
   DWF Python Example
   Author:  Digilent, Inc.
   Revision:  2018-07-28

   Requires:                       
       Python 2.7, 3
"""

from ctypes import *
from lib.AnalogDiscovery2.dwfconstants import *
import math
import time
import sys
import numpy as np
import matplotlib.pyplot as plt
from tqdm import tqdm, trange


def AnalogImpedance_Analyzer(steps = 101, start = 1e3, stop = 1e6, reference = 10e3, amplitude = 1, offset = 0, Probe_capacitance = 0, Probe_resistance = 10e6, averaging = 1):
    steps = int(steps)
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
    # ohmRes = 10e6 # Probe resistance
    # faradCap = 195e-12 # Probe capacitance
    probe_gain = 10
    print("Reference: "+ ''.join(map(str, num2SIunit(reference)))+" Ohm  Frequency: "+''.join(map(str, num2SIunit(start)))+" Hz to "+''.join(map(str, num2SIunit(stop)))+" Hz")
    dwf.FDwfAnalogImpedanceReset(hdwf)
    # see https://digilent.com/reference/test-and-measurement/guides/waveforms-impedance-analyzer for setup diagram
    dwf.FDwfAnalogImpedanceModeSet(hdwf, c_int(0)) # 0 = W1-C1-DUT-C2-R-GND, 1 = W1-C1-R-C2-DUT-GND, 8 = AD IA adapter
    dwf.FDwfAnalogImpedanceReferenceSet(hdwf, c_double(reference)) # reference resistor value in Ohms
    dwf.FDwfAnalogImpedanceFrequencySet(hdwf, c_double(start)) # frequency in Hertz
    dwf.FDwfAnalogImpedanceAmplitudeSet(hdwf, c_double(amplitude)) # 1V amplitude = 2V peak2peak signal
    dwf.FDwfAnalogImpedanceProbeSet(hdwf, c_double(Probe_resistance), c_double(Probe_capacitance)) #  Specifies the probe impedance that will be taken in consideration for measurements
    if offset > 0 and offset < 1:
        dwf.FDwfAnalogImpedanceOffsetSet(hdwf, c_double(offset))    # Configures the stimulus signal offset
    # dwf.FDwfAnalogImpedancePeriodSet(hdwf, c_double(cMinPeriods))

    dwf.FDwfAnalogImpedanceConfigure(hdwf, c_int(1)) # start
    time.sleep(2)
    
    rg = {label: [0.0]*steps for label in ['Hz', 'Rs', 'Xs', 'Phase', 'Cs', 'Cp', 'Ls', 'Lp']}
     
    for i in trange(steps):
        hz = stop * pow(10.0, 1.0*(1.0*i/(steps-1)-1)*math.log10(stop/start)) # exponential frequency steps
        rg['Hz'][i] = hz
        dwf.FDwfAnalogImpedanceFrequencySet(hdwf, c_double(hz)) # frequency in Hertz
        time.sleep(0.01) 
        dwf.FDwfAnalogImpedanceStatus(hdwf, None) # ignore last capture since we changed the frequency
        for _ in range(averaging):
            while True:
                if dwf.FDwfAnalogImpedanceStatus(hdwf, byref(sts)) == 0:
                    dwf.FDwfGetLastErrorMsg(szerr)
                    print(str(szerr.value))
                    quit()
                if sts.value == 2:
                    break
            resistance = c_double()
            reactance = c_double()
            phase = c_double()
            SeriesCapactance = c_double()
            ParallelCapactance = c_double()
            SeriesInductance = c_double()
            ParallelInductance = c_double()
            # measurement options in WaveForms™ SDK Reference Manual p.101
            dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceResistance, byref(resistance))
            dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceReactance, byref(reactance))
            dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceImpedancePhase, byref(phase))
            dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceSeriesCapactance, byref(SeriesCapactance))
            dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceParallelCapacitance, byref(ParallelCapactance))
            dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceSeriesInductance, byref(SeriesInductance))
            dwf.FDwfAnalogImpedanceStatusMeasure(hdwf, DwfAnalogImpedanceParallelInductance, byref(ParallelInductance))
            rg['Rs'][i] += resistance.value/averaging 
            rg['Xs'][i] += reactance.value/averaging
            rg['Phase'][i] += math.degrees(phase.value)/averaging
            rg['Cs'][i] += SeriesCapactance.value/averaging
            rg['Cp'][i] += ParallelCapactance.value/averaging
            rg['Ls'][i] += SeriesInductance.value/averaging
            rg['Lp'][i] += ParallelInductance.value/averaging



    dwf.FDwfAnalogImpedanceConfigure(hdwf, c_int(0)) # stop
    dwf.FDwfDeviceClose(hdwf)
    return rg


def num2SIunit(num):
    SIunit = {
        'p': 1e-12,
        'n': 1e-9,
        'u': 1e-6,
        'm': 1e-3,
        '': 1,
        'k': 1e3,
        'M': 1e6,
        'G': 1e9,
        'T': 1e12
    }
    for unit_, exponent_ in SIunit.items():
        if num >= exponent_:
            unit, exponent = unit_, exponent_
        else:
            return num/exponent, unit
    return num, ''

if __name__ == '__main__':
    AnalogImpedance_Analyzer()


# rgHz = [0.0]*steps
# rgRs = [0.0]*steps
# rgXs = [0.0]*steps
# rgPs = [0.0]*steps
# rgCs = [0.0]*steps
# rgCp = [0.0]*steps
# rgLs = [0.0]*steps
# rgLp = [0.0]*steps