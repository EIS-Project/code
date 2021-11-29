import board
import digitalio
import time

class led:
    def __init__(self, LED_pin=board.D5) -> None:
        self.led = digitalio.DigitalInOut(LED_pin)
        self.led.direction = digitalio.Direction.OUTPUT
        self.led.value = True

    def blink(self, x):
        try:
            x = int(x)
            for _ in range(x*2):
                self.led.value ^= True
                time.sleep(0.5)
        except:
            return False