from lib.AnalogDiscovery2.AnalogImpedance_Analyzer import AnalogImpedance_Analyzer
import logging
from tqdm import trange
import json
import os


class AnalogMux:
    def __init__(self, SerialComm, num_channels=16) -> None:
        self.num_channels = num_channels
        self.ser = SerialComm
        # self._Z_oc = self.load_Z_oc()

    @property
    def Z_oc(self):
        return self._Z_oc

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
        if os.path.exists('configuration.json'):
            with open('configuration.json') as f:
                MSMT_param = json.load(f)['MSMT_param']
        else:
            logging.info('using default MSMT_param instead of loading from json file')
            MSMT_param = {
                "steps": 501,
                "start": 1e3,
                "stop": 10e6,
                "reference": 0.989e6,
                "amplitude": 100e-3,
                "offset": 200e-3,
                "Probe_resistance": 1e6,
                "Probe_capacitance": 375e-12
            }
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
