import pandas as pd
from lib.AnalogDiscovery2.AnalogImpedance_Analyzer import AnalogImpedance_Analyzer
import logging
from tqdm import trange
import json
import os
import numpy as np


class AnalogMux:
    def __init__(self, SerialComm, num_channels=16) -> None:
        self.num_channels = num_channels
        self.ser = SerialComm
        self.C_load = 0
        self.Ron = 4
        self.Cd = 175e-12
        self.R_load = 820e3
        self.t_on = 0
        self.t_off = 0
        # self._Z_oc = self.load_Z_oc()

    @property
    def Z_oc(self):
        return self._Z_oc

    def tsett(self, on):
        """calculate settling time of the switch
        where Settling time is the time required for the switch output
        to settle within a given error band of the final value

        Args:
            on (bool): True for switching on, False for switching on
        """

        # table for Number of Time Constants Required to Settle to 1 LSB Accuracy for a Single-Pole System
        table = pd.DataFrame({'Resolution': [6, 8, 10, 12, 14, 16, 18, 20, 22],
                              'LSB': [1.563, 0.391, 0.0977, 0.0244, 0.0061, 0.00153, 0.00038, 0.000095, 0.000024],
                              'Time Constants': [4.16, 5.55, 6.93, 8.32, 9.7, 11.09, 12.48, 13.86, 15.25]})
        error = 0.1     # % error
        if on:
            return self.t_on + (self.Ron * self.R_load / (self.Ron + self.R_load)) * (self.C_load+self.Cd) * (-np.log(error/100))
        else:
            return self.t_off + self.R_load * (self.C_load+self.Cd) * (-np.log(error/100))

    def load_Z_oc(self):
        """load previous open circuit impedance calibration if the calibration file exists, 
        else run the open circuit impedance calibration and save as OpenCktImpedCalib.json in the current directory.

        Returns:
            [list of dicts]: [calibration result]
        """
        if os.path.exists('OpenCktImpedCalib.json'):
            with open('OpenCktImpedCalib.json', 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            input(
                'please make sure pogos are not in contact with the board, press enter when ready')
            with open('OpenCktImpedCalib.json', 'w', encoding='utf-8') as f:
                json.dump(self.Open_Circuit_Impedance(),
                          f, ensure_ascii=False, indent=4)

    def Open_Circuit_Impedance(self):
        """measure the open circuit imepdance for compensation

        Returns:
            [list of dicts]: [parameters of open circuit imepdance measurement]
        """
        if os.path.exists('configuration.py'):
            from configuration import configuration
            MSMT_param = configuration['MSMT_param']
        else:
            raise FileNotFoundError('configuration.py not found')
        Z_oc = [None]*self.num_channels
        for i in trange(self.num_channels):
            self.switch_channel(i+1)
            Z_oc[i] = AnalogImpedance_Analyzer(**MSMT_param, averaging=5)
            logging.info(
                f'Complete open ckt impedance calibration of channel {i+1}')
        self.switch_channel(0)      # disconnect all channels

        return Z_oc

    def switch_channel(self, channel):
        """switch adg725 analog mux to the specified channel, the analog mux disconnect all channels if channel < 1

        Args:
            channel ([int]): [channel number]

        Raises:
            ConnectionError: [if channel switches unsuccessfully]
        """
        msg = self.ser.RW(channel)
        # manual connect
        # msg = serial_read_write(key, port='COM7')
        if 'off' or 'ok' in msg:
            print(msg)
        else:
            logging.error(
                'usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
            raise ConnectionError(
                'usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
