""" 
AD725 analog mux switch using SPI
graphical datasheet: https://cdn.sparkfun.com/assets/learn_tutorials/4/5/4/graphicalDatasheet-Dev.pdf
"""
from lib.ADG725 import ADG725
from lib.LED import led
import supervisor

def main():
    led()   # initialize led
    mux = ADG725()  # initialize switch
    print('ready')
    
    while True:
        if supervisor.runtime.serial_bytes_available:   # if received data from serial bus
            data = int(input('').strip())
            if data <= 0:
                mux.disable()
                print('all switches turned off')
            else: 
                mux.switch_channel(data)   # get channel no. and switch channel, return true for success, false for fail
                print('ok, switched to channel ', data)
                # blink(data, led)

            

if __name__ == '__main__':
    main()


