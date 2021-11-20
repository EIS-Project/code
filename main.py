"""setup and measure data from Analog Discovery Kit 2.
"""
from AnalogImpedance_Analyzer import AnalogImpedance_Analyzer
from AnalogImpedance_Compensation import AnalogImpedance_Compensation

import pandas as pd
import os
import logging
import win32com.client
from contextlib import suppress
from pathlib import Path


import serial
import serial.tools.list_ports
import time
from datetime import datetime
from tqdm import trange
from Summary import summary
from DUT import DUT

def main():
    DUT_info = {
        # DUT no.: info
        1: '75kOhm',
        2: '820kOhm',
        #6: 'W1DUT1M'
        #12: '198.2KOhm',
        14: '10kOhm',
        16: '1.003kOhm'
    }
        
    main_path = r"C:\Users\User01\Desktop\EIS\Cal Poly\EIS Team - MSMT_data"
    AD2_capacitance = 200e-12 # capacitance of Analog Discovery 2 + Probes
    ADG725_capacitance = 175e-12 # Cd of ADG725 
    MSMT_param = {
        'steps': 501,
        'start': 1e3,     
        'stop': 10e6,
        'reference': 0.989e6, #0.995e6,
        'amplitude': 100e-3,
        'offset': 200e-3,
        'Probe_resistance': 1e6,
        'Probe_capacitance': AD2_capacitance+ADG725_capacitance
    }
    DUTs = {}
    interval = 10  # interval between measurements in seconds
    total_seconds = 60
    start_time = time.time()
    current_time = time.time()
    log_folder = os.path.join(main_path, 'log_file', f'{datetime.now():%m_%d_%Y}')
    Path(log_folder).mkdir(parents=True, exist_ok=True)
    log_file = os.path.join(log_folder, f'{datetime.now():%m_%d_%H_%M_%Y}.log')
    shortcut = os.path.join(main_path, 'measurement status.lnk')
    logging.basicConfig(format='%(asctime)s - %(message)s', filename = log_file, level=logging.INFO)
    create_shortcut(log_file, shortcut)
    logging.info('Program begins')
    experiment_date = f'{datetime.now():%m_%d_%Y_%H_%M}'
    while current_time - start_time < total_seconds:
        logging.info(f'measurment time left: {total_seconds - current_time + start_time}s')
        print(current_time - start_time)
        for key, val in DUT_info.items():
            date = f'{datetime.now():%m/%d/%Y %H:%M:%S}'
            DUT_folder = os.path.join(main_path, f'{val}')
            test_result_folder = os.path.join(main_path, f'{val}', experiment_date)
            if key not in DUTs:
                DUTs[key] = DUT()
                DUTs[key].create_folder(DUT_folder)     # create folder for each DUT
                DUTs[key].create_folder(test_result_folder)  # create sub folder to store individual test result
                DUTs[key].DUT_info = DUT_info[key]  # store the info of the DUT
            
            DUTs[key].writer = pd.ExcelWriter(os.path.join(test_result_folder, f'{datetime.now():%m_%d_%Y_%H_%M}.xlsx'))  # create excel file for individual test under subfolder
            DUTs[key].sheetname = f'{datetime.now():%m_%d_%Y_%H_%M}'
            ## control ADG725 to change DUT channel
            msg = auto_connect(key)
            ## manual connect
            # msg = serial_read_write(key, port='COM7')
            if 'ok' not in msg:
                logging.error('usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
                return
            print(msg)
            logging.info(f'Begin impedance measurement of {val}')
            data = AnalogImpedance_Analyzer(**MSMT_param)
            DUTs[key].Generate_Report(data, date, current_time - start_time, **MSMT_param)
            DUTs[key].writer.save()
            logging.info(f'Finished impedance measurement of {val}')
        print('waiting for next measurement')
        ## delay with progress bar
        for _ in trange(interval):
            time.sleep(1)
        current_time = time.time()
        
    
    ## generate Summary file for each DUT
    for val in DUT_info.values():
        logging.info(f'Generate Summary file for {val}')
        test_result_folder = os.path.join(main_path, f'{val}', experiment_date)
        summary(test_result_folder)
    logging.info('Program Finished')

    ## remove shortcut
    with suppress(FileNotFoundError):
        os.remove(shortcut)
    

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


def auto_connect(msg):
    # scan all the ports
    for port in sorted([port[0] for port in list(serial.tools.list_ports.comports())], key=lambda x:int(x.split('COM')[-1]), reverse=False):
        print(port)
        try:
            rx = serial_read_write(msg, port)
            if 'ok' in rx:      # wait for acknowledgement,  device returns ok if the correct PORT is connected
                return rx
            else:
                logging.error('usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
                return None
        except serial.serialutil.SerialException:   # else try to conenct next port
            pass
    return 


def serial_read_write(msg, port='COM7'):
    """send serial data to MCU

    Args:
        msg (str): msg to MCU, DUT number in this case
        port (str, optional): COM port number. Defaults to 'COM7'.

    Returns:
        str: msg reply from MCU, currently set as 'ok, switched to DUT {msg}'
    """
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
    
def create_shortcut(root_file, output_path):
    '''create shortcut of the root file
    output_path: path of the shortcut file
    root_file: location of the root file
    '''
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(output_path)
    shortcut.Targetpath = root_file
    shortcut.IconLocation = root_file
    shortcut.save()


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logging.error(e)


# def write_read(msg, arduino):
#     arduino.write(bytes(msg, 'utf-8'))
#     time.sleep(0.1)
#     return arduino.readline(arduino.inWaiting()).strip().decode('utf-8')

# def summary(output_dir, input_dir):
#     """generate summary file, overlay all the results in one plot
#     output_dir: where the summary file is saved
#     input_dir: where the data files are
#     """
#     Summary = pd.ExcelWriter(os.path.join(output_dir, 'Summary.xlsx'))
#     summary_table = []
#     for root, dirs, files in os.walk(input_dir):
#         for file in files:
#             summary_table.append(pd.read_excel(os.path.join(input_dir,file), nrows=1, usecols=list(range(9,21))))
#         df = pd.concat(summary_table)
#         time_elapsed = df['Time Elapsed [s]'].to_list()
#         df['date'] = pd.to_datetime(df['date'])  # convert date column in datetime format
#         df.sort_values(by=['date'], ascending=False, inplace=True)
#         print(df)
#         df.to_excel(Summary, f'Summary', index=False,)
#         worksheet = Summary.sheets['Summary']
#         # Link to another Excel workbook.
#         for row, file in enumerate(files):
#             worksheet.write_url(row+1, df.shape[1]-1, os.path.join(input_dir,file), string=str(time_elapsed[row]))
#         worksheet.set_column(0, df.shape[1]-1, 17.5)
#         Summary.save()