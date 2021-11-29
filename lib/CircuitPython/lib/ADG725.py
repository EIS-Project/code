import board
import busio
import digitalio

class ADG725:
    def __init__(self, CS=board.D10, CLK=board.D13, MOSI=board.D11, MISO=None) -> None:
        self.CLK = CLK
        self.MOSI = MOSI
        self.CS = digitalio.DigitalInOut(CS)
        self.CS.switch_to_output(value=True)
        # cs.direction = digitalio.Direction.OUTPUT
        self.CS.value = True
        self._EN = True
    

    def switch_channel(self, channel):
        """switch channel

        Args:
            channel ([int]): [channel number]
        """
        EN = 7
        CSA = 6
        CSB = 5
        DATA = 0x00
        if self._EN:
            DATA &= ~( 1 << EN | 1 << CSA | 1 << CSB)
            DATA |= ((channel-1) & 0x0F)
        else:
            DATA |= (1 << EN)   # turn off switches
        buf = bytearray(1)
        buf[0] = DATA

        with busio.SPI(clock=self.CLK, MOSI=self.MOSI) as spi:
            while not spi.try_lock():
                pass
            spi.configure(baudrate=5000000, phase=1, polarity=0)
            self.CS.value = False
            spi.write(buf)
            self.CS.value = True
            return True
    
    def disable(self):
        """turn off all the switches
        """
        self._EN = False
        self.switch_channel()
        self._EN = True
