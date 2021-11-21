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
import time
from datetime import datetime
from tqdm import trange
from Summary import summary
from DUT import DUT
from SerialComm import SerialComm

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
    ## setup serial communication with microcontroller
    serial1 = SerialComm(auto_connect=True)

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
            msg = serial1.RW(key)
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
