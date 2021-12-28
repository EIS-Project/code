configuration = {
    # information of dut for each channel
    "DUT_info": {
        # DUT no.: info
        1: "75kOhm",
        2: "820kOhm",
        # 6: "W1DUT1M"
        # 12: "198.2KOhm",
        14: "10kOhm",
        16: "1.003kOhm"
    },
    # impedance analyzer measurement settings
    "MSMT_param": {
        "steps": 501,
        "start": 1e3,
        "stop": 10e6,
        "reference": 0.989e6,
        "amplitude": 100e-3,
        "offset": 200e-3,
        "Probe_resistance": 1e6,
        # Probe_capacitance = capacitance of Analog Discovery 2 (170e-12) + ADG725 drain capacitance (175e-12) + probe capacitance (30e-12)
        "Probe_capacitance": 375e-12
    },
    # Arduino serial communication settings
    "Arduino": {
        "COM": "COM7",
        "baudrate": 9600
    }


}
