from lib.AnalogDiscovery2.AnalogImpedance_Analyzer import AnalogImpedance_Analyzer
import logging
from tqdm import trange

class AnalogMux:
    def __init__(self, MSMT_param, ser, num_channels=16) -> None:
        self.num_channels = num_channels
        self.Z_oc = self.Open_Circuit_Impedance()
        self.MSMT_param = MSMT_param
        self.ser = ser
    
    def Open_Circuit_Impedance(self):
        """measure the open circuit imepdance for compensation

        Returns:
            [list of dicts]: [parameters of open circuit imepdance measurement]
        """
        Z_oc = [None]*self.num_channels
        for i in trange(self.num_channels):
            self.switch_channel(i)
            Z_oc[i] = AnalogImpedance_Analyzer(**self.MSMT_param, averaging=5)
        return Z_oc
            
    
    def switch_channel(self, channel):
        msg = self.ser.RW(channel)
        ## manual connect
        # msg = serial_read_write(key, port='COM7')
        if 'ok' not in msg:
            logging.error('usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
            raise ConnectionError('usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
        print(msg)