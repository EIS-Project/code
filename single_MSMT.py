"""setup and measure data from Analog Discovery Kit 2.
"""
from AnalogImpedance_Analyzer import AnalogImpedance_Analyzer
from AnalogImpedance_Compensation import AnalogImpedance_Compensation

import pandas as pd
import os
import numpy as np
from pathlib import Path

from scipy.signal import savgol_filter
import serial
import serial.tools.list_ports
import time


def main():

    DUTs = {
        # key = DUT no.
        # value = info
        1: '10kOhm',
        # 2: '820kOhm',
        #6: 'W1DUT1M'
        #12: '198.2KOhm',
        # 14: '10kOhm',
        # 16: '1.003kOhm'
    }
        


    steps = 501
    start = 1e3     
    stop = 10e6
    reference = 0.995e6
    amplitude = 100e-3
    offset = 0e-3
    Probe_resistance = 1e6
    Probe_capacitance = 100e-12
    GR = Generate_Report()
    for key, val in DUTs.items():
        folder = os.path.join(r"C:\Users\User01\Desktop\EIS\Impedance_measurement\MSMT_data", f'{val}')
        Path(folder).mkdir(parents=True, exist_ok=True)    # create folder if not exist
        # msg = serial_read_write(key, port='COM7')
        msg = 'ok'
        time.sleep(0.1)
        start_time = time.time()
        while 'ok' not in msg:      # wait for acknowledgement
            if time.time() - start_time > 1:
                print('usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
                return
        print(msg)
        data = AnalogImpedance_Analyzer(steps, start, stop, reference, offset, Probe_capacitance, Probe_resistance)
        GR.sheetname = f'DUT{key}'
        GR.Generate_Report(data, start, stop, steps, amplitude, DUTs[key], reference, offset, False, Probe_capacitance, Probe_resistance)
    GR.writer.save()
    
class Generate_Report():
    def __init__(self):
        self.writer = None
        self.sheetname = None
        self.workbook = None

    def Generate_Report(self, data, start, stop, steps, Amplitude, DUT_info, reference, offset, filter=True, Probe_capacitance = 0, Probe_resistance = 0):
        info = pd.DataFrame({
            'Amplitude': Amplitude,
            'Reference [Ohm]': reference,
            'DUT Info': DUT_info,
            'Offset [V]': offset,
            'Probe resistance [Ohm]': Probe_resistance, 
            'Probe Cap [pF]': Probe_capacitance*1e12 
        }, index=[0])
        
        steps = int(steps)

        df = pd.DataFrame(data)
        # add info of DUT to current dataframe
        
        df.insert(1, '|Z|', (df['Rs'].pow(2)+df['Xs'].pow(2)).pow(1./2)) 
        df['Cs'] *= 1e6
        df['Cp'] *= 1e9
        if filter:
            for col in df.columns:
                if col != 'Hz':
                    window = ~(steps//7 % 2) + steps//7
                    df[col] = savgol_filter(df[col], window, 9)
        filename = 'Z={}+j{}'.format(str(int(np.mean(data['Rs']))), str(int(np.mean(data['Xs']))))
        if self.writer is None:
            self.writer = pd.ExcelWriter(os.path.join(os.path.dirname(__file__), 'MSMT_data', filename+'.xlsx'))
            self.workbook = self.writer.book
        # sheetname = 'Amplitude='+''.join(map(str, num2SIunit(amplitude)))+'V'
        chart_colpos = df.shape[1]
        df = pd.concat([df, info], axis=1)
        df.to_excel(self.writer, self.sheetname, index=False)
        
        worksheet = self.writer.sheets[self.sheetname]
        for col, label in enumerate(['Impedance [kOhm]', 'Resistance [kOhm]', 'Reactance [kOhm]', 'Phase [degree]', 'Series Capaitance [uF]', 'Parallel Capaitance [nF]', 'Series Inductance [H]', 'Parallel Inductance [H]']):
            chart = self.workbook.add_chart({'type': 'scatter'})
            chart.add_series({
                'name': label,
                'categories': [self.sheetname, 1, 0, df.shape[0], 0],  # x-axis
                'values': [self.sheetname, 1, col+1, df.shape[0], col+1],    # y-axis
                'line': {'none': True},
                'marker': {'type': 'circle', 'size': 3}
            })
            chart.set_x_axis({
                'name': 'Frequency [kHz]',
                'major_gridline': {'visible': True},
                'log_base': 10,
                'min': start, 'max': stop,
                'display_units': 'thousands', 
                'display_units_visible': False,
                'minor_tick_mark': 'cross'})
            if label == 'Phase [degree]':
                chart.set_y_axis({'name': label, 'major_gridline': {'visible': True}, 'major_unit': 30})
            elif 'kOhm' in label:
                chart.set_y_axis({
                    'name': label,
                    'major_gridline': {'visible': True},
                    'display_units': 'thousands', 
                    'display_units_visible': False,
                    })
            else:
                chart.set_y_axis({'name': label, 'major_gridline': {'visible': True}})
            chart.set_title({'name': label})
            chart.set_legend({'position': 'bottom'})
            chart.set_size({'x_scale': 1.2, 'y_scale': 1.5})
            worksheet.insert_chart(3+col*22, chart_colpos, chart)        # df.shape[0]: len(df.index), df.shape[1]: len(df.columns)




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
        if abs(num) >= exponent_:
            unit, exponent = unit_, exponent_
        else:
            return num/exponent, unit
    return num, ''

def IQR_filter(df):
    # IQR filtering
    cols = []
    for col in df.columns:
        Q3 = df[col].quantile(0.75)
        Q1 = df[col].quantile(0.25)
        IQR = Q3 - Q1
        cols += [df[col].to_frame().query(f'(@Q1 - 1.5 * @IQR) <= {col} <= (@Q3 + 1.5 * @IQR)')]
    return pd.concat(cols, axis=1)

def moving_avg(df, window=7):
    for col in df.columns:
        if col != 'Hz':
            df[col] = df[col].rolling(window=window).mean()
    return df


def scan_ports(target_port):
    ports = list(serial.tools.list_ports.comports())
    for p in ports:
        if target_port in p.description:
            return p[0]

def serial_read_write(msg, port='COM7'):
    with serial.Serial(port=port, baudrate=9600, timeout=0.01) as ser:
        ser.write(f"{msg}\r".encode('utf-8'))
        time.sleep(0.1)
        ser.flushOutput()
        return read_data(ser)

def read_data(ser):
    data = ser.read(1)
    data+= ser.read(ser.inWaiting())
    if data:
        return data.strip().decode('utf-8')

def ckeck_num(data, minval, maxval):
    """check if data is within minval and maxval

    Args:
        data (str | int): input data
        minval (int): minimum value
        maxval (int): maximam value

    Raises:
        ValueError: if data not within range or not number

    Returns:
        num: input data within range
    """
    try:
        num = int(data)
        if num in range(minval,maxval+1):
            return num
        raise ValueError
    except ValueError:
        print('invalid input: ', end='')
        return ckeck_num(input(f'enter number between {minval}-{maxval}: '), minval, maxval)
    


if __name__ == '__main__':
    main()


# def write_read(msg, arduino):
#     arduino.write(bytes(msg, 'utf-8'))
#     time.sleep(0.1)
#     return arduino.readline(arduino.inWaiting()).strip().decode('utf-8')

