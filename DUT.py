import pandas as pd
from pathlib import Path
from scipy.signal import savgol_filter
from lib.AnalogDiscovery2.AnalogImpedance_Analyzer import AnalogImpedance_Analyzer
from datetime import datetime
import os
import time
from lib.generate_result.Summary import summary
import logging

class DUT:
    def __init__(self, MSMT_param, channel, DUT_info, main_path, ser, mux):
        self.writer = None
        self.workbook = None
        self.window_size = 7  # set window size for filtering impedance measurement using Savitzky–Golay filter
        self.DUT_info = DUT_info        # store the info of the DUT
        self._channel = channel
        self.MSMT_param = MSMT_param
        self.main_path = main_path
        experiment_date = f'{datetime.now():%m_%d_%Y_%H_%M}'
        self.DUT_folder = os.path.join(self.main_path, f'{self.DUT_info}')
        self.test_result_folder = os.path.join(self.main_path, f'{self.DUT_info}', experiment_date)
        self.create_folder(self.DUT_folder)     # create folder for each DUT
        self.create_folder(self.test_result_folder)  # create sub folder to store individual test result
        self.data = None
        self.mux = mux
        # self.Z_oc = self.mux.Z_oc[self.channel]    # open circuit impedance of the current channel
        self.start_time = time.time()
        
        
        
    @property
    def channel(self):
        return self._channel
    

    def create_folder(self, dirpath):
        """create folder if not exist"""
        Path(dirpath).mkdir(parents=True, exist_ok=True)

    def Generate_Report(self):
        self._Generate_Report(**self.MSMT_param) # too lazy to replace parameters in _Generate_Report function one by one

    def _Generate_Report(self, start, stop, steps, amplitude, reference, offset, Probe_capacitance = 0, Probe_resistance = 0):
        """generate report from the impedance measuremets from a single frequency sweep

        Args:
            start ([float]): [start frequency]
            stop ([float]): [stop frequency]
            steps ([int]): [frequency step size]
            amplitude ([float]): [amplitude of the applied signal]
            reference ([float]): [reference resistor value]
            offset ([float]): [offset voltage of the applied signal]
            Probe_capacitance (int, optional): [parasitic capacitance from the measurement point of the DUT to the input of the system]. Defaults to 0.
            Probe_resistance (int, optional): [parasitic resistance from the measurement point of the DUT to the input of the system]. Defaults to 0.
        """
        elapsed = time.time() - self.start_time
        data = self.data
        date = f'{datetime.now():%m/%d/%Y %H:%M:%S}'
        self.writer = pd.ExcelWriter(os.path.join(self.test_result_folder, f'{datetime.now():%m_%d_%Y_%H_%M}.xlsx'))  # create excel file for individual test under subfolder
        self.sheetname = f'{datetime.now():%m_%d_%Y_%H_%M}'
        
        info = pd.DataFrame({
            'date': date,
            'Connection': 'BreadBoard',
            'TestBoard': 'POGO',
            'Compensation [pF]': Probe_capacitance*1e12,
            'Start Freq [Hz]': start,
            'Stop Freq [Hz]': stop,
            'Rref [Ohm]': reference,
            'Amplitude': amplitude,
            'Offest [V]': offset,
            'Samples': steps,
            'DUT Info': self.DUT_info,
            'Time Elapsed [s]': elapsed
        }, index=[0])
        
        steps = int(steps)

        df = pd.DataFrame(data)
        # add info of DUT to current dataframe
        
        df.insert(1, '|Z|', (df['Rs'].pow(2)+df['Xs'].pow(2)).pow(1./2)) 
        df['Cs'] *= 1e6
        df['Cp'] *= 1e9
        if self.window_size > 1:
            for col in df.columns:
                if col != 'Hz':
                    window = ~(steps//self.window_size % 2) + steps//self.window_size
                    df[col] = savgol_filter(df[col], window, 9)
        # filename = 'Z={}+j{}'.format(str(int(np.mean(data['Rs']))), str(int(np.mean(data['Xs']))))

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
        self.writer.save()
    
    def Generate_Summary(self):
        """call summary function to generate summary file from all the measurements of the current DUT
        """
        logging.info(f'Generate Summary file for {self.DUT_info}')
        summary(self.test_result_folder)



    def IQR_filter(self, df):
        # IQR filtering
        cols = []
        for col in df.columns:
            Q3 = df[col].quantile(0.75)
            Q1 = df[col].quantile(0.25)
            IQR = Q3 - Q1
            cols += [df[col].to_frame().query(f'(@Q1 - 1.5 * @IQR) <= {col} <= (@Q3 + 1.5 * @IQR)')]
        return pd.concat(cols, axis=1)

    def Measure_Impedance(self):
        logging.info(f'Begin impedance measurement of {self.DUT_info}')
        self.data = AnalogImpedance_Analyzer(**self.MSMT_param)
        logging.info(f'Finished impedance measurement of {self.DUT_info}')

    def moving_avg(self, df, window=7):
        for col in df.columns:
            if col != 'Hz':
                df[col] = df[col].rolling(window=window).mean()
        return df
 
    def num2SIunit(self, num):
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
