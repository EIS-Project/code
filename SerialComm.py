import serial
import time
import logging

class SerialComm:
    # serial communication class
    def __init__(self, COM='COM7', auto_connect=False) -> None:
        self.ser = None
        if auto_connect:
            self.auto_connect(1)
        else:
            self.COM = COM  # get COM port manually


    def __enter__(self):
        if not self.COM:
            self.auto_connect(1)


    def __exit__(self, type, value, traceback):
        if self.ser:
            self.ser.close()

    def auto_connect(self, msg):
        # scan all the ports
        sorted_ports = sorted([port[0] for port in list(serial.tools.list_ports.comports())], key=lambda x:int(x.split('COM')[-1]), reverse=False)  # ports sorted in ascending order
        for self.COM in sorted_ports:
            print(f'connecting to port{self.COM}')
            try:
                rx = self.RW(msg)
                if 'ok' in rx:      # wait for acknowledgement,  device returns ok if the correct PORT is connected
                    return rx
                else:
                    logging.error('usb timeout, program terminated, try clicking reset bottom on the PCB to resolve the issue')
                    return None
            except serial.serialutil.SerialException:   # else try to conenct next port
                pass
        return 


    def RW(self, msg):
        """send serial data to MCU

        Args:
            msg (str): msg to MCU, DUT number in this case
            port (str, optional): COM port number. Defaults to 'COM7'.

        Returns:
            str: msg reply from MCU, currently set as 'ok, switched to DUT {msg}'
        """
        with serial.Serial(port=self.COM, baudrate=9600, timeout=0.01) as self.ser:
            self.ser.write(f"{msg}\r".encode('utf-8'))
            time.sleep(0.1)
            self.ser.flushOutput()
            return self.read_data()

    def read_data(self):
        data = self.ser.read(1)
        data+= self.ser.read(self.ser.inWaiting())
        if data:
            return data.strip().decode('utf-8')
