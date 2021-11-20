import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd
import os

filetypes = (
        ('excel csv', '*.csv'),
        ('excel xlsx', '*.xlsx'),
        ('All files', '*.*')
    )

files = fd.askopenfilenames(title='Open a file',
        initialdir=os.path.dirname(__file__),filetypes=filetypes)
writer = pd.ExcelWriter(os.path.join(os.path.dirname(__file__), 'MSMT_data', 'Digilent_WaveForms_Network_Analyzer'+'.xlsx'))
for file in files:
    print(file)
    
    if file.endswith('csv'):
        df = pd.read_csv(file)
        filename = os.path.basename(file).replace('.csv', '')
    elif file.endswith('xlsx'):
        df = pd.read_excel(file)
        filename = os.path.basename(file).replace('.xlsx', '')
    
    df.to_excel(writer, sheet_name=filename, index=False)
    workbook = writer.book
    worksheet = writer.sheets[filename]
    chart = workbook.add_chart({'type': 'scatter'})
    
    chart.add_series({
        'name': df.columns[1],
        'categories': [filename, 1, 0, df.shape[0], 0],  # x-axis
        'values': [filename, 1, 1, df.shape[0], 1],    # y-axis
        'line': {'none': True},
        'marker': {'type': 'circle', 'size': 3},
        'y2_axis': True
    })
    chart.add_series({
        'name': df.columns[2],
        'categories': [filename, 1, 0, df.shape[0], 0],  # x-axis
        'values': [filename, 1, 2, df.shape[0], 2],    # y-axis
        'line': {'none': True},
        'marker': {'type': 'circle', 'size': 3}
    })
    chart.set_x_axis({
        'name': 'Frequency [kHz]',
        'label_position': 'low',
        'log_base': 10,
        'display_units': 'thousands',
        'display_units_visible': False,
        'major_gridlines': {'visible': True},
        'major_gridlines': {'visible': False},
        'minor_tick_mark': 'none',
        'min': df[df.columns[0]].min()
    })
    chart.set_y_axis({
        'name': df.columns[2],
        'major_gridlines': {'visible': True}
    })
    chart.set_y2_axis({
    'name': df.columns[1],
    'major_gridlines': {'visible': True}
    })
    chart.set_title({'name': filename})
    chart.set_legend({'position': 'bottom'})
    chart.set_size({'x_scale': 1.2, 'y_scale': 1.5})
    worksheet.insert_chart(1, df.shape[1]+4, chart)
writer.save()