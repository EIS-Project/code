from os import write
from AnalogIn_Logger import AnalogIn_Logger
import pandas as pd
import numpy as np
import os

freq = 100e3
amplitude = 100e-3
data = AnalogIn_Logger(frequency=freq, amplitude=amplitude)
data = np.array(data)
df = pd.DataFrame(data.T)
columns = ['Channel 1', 'Channel 2']
df.columns = columns
df.insert(0, 'time [s]', df.index/freq)

filename = f'{freq/1e3}kHz{amplitude}V'
writer = pd.ExcelWriter(os.path.join(os.path.dirname(__file__), 'MSMT_data', filename + '.xlsx'))
workbook = writer.book
sheetname = filename
df.to_excel(writer, sheetname, index=False)
worksheet = writer.sheets[sheetname]
chart = workbook.add_chart({'type': 'scatter'})
for col, label in enumerate(columns):
    
    chart.add_series({
        'name': label,
        'categories': [sheetname, 1, 0, df.shape[0], 0],  # x-axis
        'values': [sheetname, 1, col+1, df.shape[0], col+1],    # y-axis
        'line': {'none': True},
        'marker': {'type': 'circle', 'size': 3}
    })
chart.set_x_axis({
    'name': 'time [s]',
    'major_gridline': {'visible': True},
    'display_units': 'thousands', 
    'display_units_visible': False,
    'minor_tick_mark': 'cross'})

chart.set_y_axis({'name': label, 'major_gridline': {'visible': True}})
chart.set_title({'name': label})
chart.set_legend({'position': 'bottom'})
chart.set_size({'x_scale': 1.2, 'y_scale': 1.5})
worksheet.insert_chart(1, df.shape[1], chart)        # df.shape[0]: len(df.index), df.shape[1]: len(df.columns)
writer.save()