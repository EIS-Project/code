from sys import exc_info
import serial.tools.list_ports
import time
import logging
import os


class SerialComm:
    # serial communication class
    def __init__(self, auto_connect=False) -> None:
        if os.path.exists('configuration.py'):
            from configuration import configuration
            # load settings
            config = configuration['Arduino']
        else:
            logging.info(
                'configuration.py not found, using default configuration for MSMT_param')
            config = {
                "COM": "COM7",
                "baudrate": 9600
            }
        self.baudrate = config['baudrate']
        self.ser = None
        if auto_connect:
            self.auto_connect(1)
        else:
            self.COM = config['COM']  # get COM port manually

    def __enter__(self):
        if not self.COM:
            self.auto_connect(1)

    def __exit__(self, type, value, traceback):
        if self.ser:
            self.ser.close()

    def auto_connect(self, msg):
        # scan all the ports
        ports = serial.tools.list_ports.comports()
        # sorted_ports = sorted(ports) # sorted([port[0] for port in list(serial.tools.list_ports.comports())], key=lambda x:int(x.split('COM')[-1]), reverse=False)  # ports sorted in ascending order
        try:
            for port in sorted(ports):
                if port.pid:
                    print(f'connecting to port {port}')
                    self.COM = port.name
                    rx = self.RW(msg)
                    print(rx)
                    if 'ok' in rx:      # wait for acknowledgement,  device returns ok if the correct PORT is connected
                        return rx

            raise ConnectionError('serial communication device not found')
        except:
            logging.error(
                'usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
            logging.exception('message')

    def RW(self, msg):
        """send serial data to MCU

        Args:
            msg (str): msg to MCU, DUT number in this case
            port (str, optional): COM port number. Defaults to 'COM7'.

        Returns:
            str: msg reply from MCU, currently set as 'ok, switched to DUT {msg}'
        """
        with serial.Serial(port=self.COM, baudrate=self.baudrate, timeout=0.01) as self.ser:
            self.ser.write(f"{msg}\r".encode('utf-8'))
            time.sleep(0.1)
            self.ser.flushOutput()
            return self.read_data()

    def read_data(self):
        data = self.ser.read(1)
        data += self.ser.read(self.ser.inWaiting())
        if data:
            return data.strip().decode('utf-8')


if __name__ == '__main__':
    ser = SerialComm(auto_connect=False)
