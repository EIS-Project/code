""" 
AD725 analog mux switch using SPI
graphical datasheet: https://cdn.sparkfun.com/assets/learn_tutorials/4/5/4/graphicalDatasheet-Dev.pdf
"""
import board
import busio
import digitalio
import time
import supervisor
import microcontroller

def main():

    cs = digitalio.DigitalInOut(board.D10)
    cs.switch_to_output(value=True)
    # cs.direction = digitalio.Direction.OUTPUT
    cs.value = True

    spi = busio.SPI(clock=board.D13, MOSI=board.D11)
    
    while not spi.try_lock():
        pass
    spi.configure(baudrate=5000000, phase=1, polarity=0)

    led = digitalio.DigitalInOut(board.A5)
    led.direction = digitalio.Direction.OUTPUT

    led.value = True
    time.sleep(1)
    led.value = False

    print('begin')
    while True:
        if supervisor.runtime.serial_bytes_available:
            
            
            data = input('').strip()
            print('value entered ', data)
            switch_channel(int(data), cs, spi)
            print('led off')
            
        else:
            led.value = False

def blink(x, led):
    for i in range(x*2):
        led.value ^= True
        time.sleep(0.5)

def switch_channel(channel, cs, spi):
    EN = 7
    CSA = 6
    CSB = 5
    DATA = 0x00
    DATA &= ~( 1 << EN | 1 << CSA | 1 << CSB)
    DATA |= ((channel-1) & 0x0F)
    buf = bytearray(1)
    buf[0] = DATA
    print(buf)
    cs.value = False
    spi.write(buf)
    cs.value = True
if __name__ == '__main__':
    main()


