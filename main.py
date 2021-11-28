"""setup and measure data from Analog Discovery Kit 2.
"""

import os
import logging
import win32com.client
from contextlib import suppress
from pathlib import Path
import time
from datetime import datetime
from tqdm import trange

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
    DUTs = []
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
    ## setup serial communication with microcontroller
    ser = SerialComm(auto_connect=True)

    for key, val in DUT_info.items():    # create DUT instances
        DUTs.append(DUT(**MSMT_param, key, val, main_path, ser))

    while current_time - start_time < total_seconds:
        logging.info(f'measurment time left: {total_seconds - current_time + start_time}s')
        print(current_time - start_time)
        for _DUT in DUTs:
            _DUT.switch_channel()
            _DUT.Measure_Impedance()
            _DUT.Generate_Report()
        print('waiting for next measurement')
        ## delay with progress bar
        for _ in trange(interval):
            time.sleep(1)
        current_time = time.time()
        
    
    ## generate Summary file for each DUT
    for _DUT in DUTs:
        _DUT.Generate_Summary()
        
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
