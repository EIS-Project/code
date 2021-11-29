"""Developer: Conrad Chan
    Not Open License
"""
import pandas as pd
import os, re
from tkinter.filedialog import askopenfilenames


def main():
    """generate bode plots from LTSpice .txt file
    """
    files = askopenfilenames(initialdir=os.path.dirname(__file__), title= "Please select a file:", filetypes= [('LTSpice .txt file', '*.txt'),])
    for file in files:
        if file:
            if file.endswith('txt'):
                # doc = codecs.open(file,'rU','UTF-16') #open for reading with "universal" type set
                df = pd.read_csv(file, sep='\t', engine='python', encoding="ISO-8859-1")
                df.columns.values[0] = 'Frequency'
                data = []
                three_dB_mag = []
                three_dB_freq = []
                Q = []
                BW = []
                cols = []
                for i in range(1, df.shape[1]):
                    col = f'{df.columns[i]} [dB]'
                    cols.append(col)
                    frame = df[df.columns[i]].apply(lambda x: re.sub('[^0123456789e.,+-]','',x)).str.split(',', expand=True)
                    frame = frame.astype(float)
                    frame.columns = [col, f'{df.columns[i]} [degree]']
                    data.append(frame)
                    A0 = frame[col].max()
                    print(A0)
                    three_dB_mag.append(A0 - 3)
                    three_dB_freq.append(df.loc[frame[col] > three_dB_mag[i-1]]['Frequency'].iloc[[0, -1]])
                    if three_dB_freq[i-1].tolist()[0]  == df.loc[0, 'Frequency']:
                        BW.append(three_dB_freq[i-1].tolist()[-1] - 0)
                    else:
                        BW.append(three_dB_freq[i-1].tolist()[-1]- three_dB_freq[i-1].tolist()[0])
                    print('3dB summary')
                    print(three_dB_mag, 'dB', three_dB_freq[i-1].tolist(), 'Hz')
                    print('BW', BW, 'Hz')
                    F0 =  df.loc[frame[col] == A0, 'Frequency'].values[0]
                    Q.append(F0/BW[i-1])
                    print('Q', Q)
                df = pd.concat([df['Frequency'], pd.concat(data, axis=1)], axis=1)
                
                # df.plot(x='Frequency', y=['Amplitude [dB]', 'phase [degree]'], kind='line', subplots=True, sharex=True, logx=True)
                writer = pd.ExcelWriter(file.replace('txt', 'xlsx'))
                df = pd.concat([df, pd.DataFrame.from_dict({'trace': cols, 'BW(Hz)': BW, 'Q': Q}, orient='columns')], axis=1,  join='outer')
                df.to_excel(writer, 'data', index=False)
                workbook = writer.book
                worksheet = writer.sheets['data']
                if len(cols):
                    overlay_charts = [workbook.add_chart({'type': 'scatter'}) for _ in range(2)]
                for col, col_name in enumerate(cols):
                    points = [None]*len(df)
                    custom_labels = [[{'delete': True}]*len(df), [{'delete': True}]*len(df)]
                    # phase = [{'delete': True}]*len(df)
                    for i, idx in enumerate(three_dB_freq[col].index.values):
                        if idx != 0 and idx != len(df)-1:
                            points[idx] = {'fill': {'color': 'red'}, 'type': 'automatic', 'data_labels': {'value': True}}
                            custom_labels[0][idx] = {'value': '{}'.format(num2SIunit(three_dB_freq[col].tolist()[i])) +'Hz\n'+ '{:.2f}'.format(three_dB_mag[col])+ 'dB'}
                            phase_val = df.loc[df['Frequency'] == three_dB_freq[col].tolist()[i], df.columns[(col+1)*2]].values[0]
                            custom_labels[1][idx] = {'value': '{}'.format(num2SIunit(three_dB_freq[col].tolist()[i])) +'Hz\n'+ '{:.2f}'.format(phase_val)+ '°'}
                    for i in range(2):
                        chart = workbook.add_chart({'type': 'scatter'})
                        chart.add_series({
                            'categories': ['data', 1, df.columns.get_loc('Frequency'), len(df), df.columns.get_loc('Frequency')],
                            'values': ['data', 1, df.columns.get_loc(col_name)+i, len(df),df.columns.get_loc(col_name)+i],
                            'marker': {'type': 'automatic', 'size': 1},
                            'line':   {'none': True},
                            'points': points,
                            'data_labels': {'value': True, 'custom': custom_labels[i], 'position': 'above'},
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
                            'min': df['Frequency'].min()
                        })
                        chart.set_y_axis({
                            'name': df.columns.values[i+1],
                            'major_gridlines': {'visible': True}
                        })
                        chart.set_legend({'none': True})
                        chart.set_title({'name': df.columns.values[i+1]})
                        
                        if len(cols) > 1:
                            overlay_charts[i].add_series({
                                'name': col_name,
                                'categories': ['data', 1, df.columns.get_loc('Frequency'), len(df), df.columns.get_loc('Frequency')],
                                'values': ['data', 1, df.columns.get_loc(col_name)+i, len(df),df.columns.get_loc(col_name)+i],
                                'marker': {'type': 'automatic', 'size': 1},
                                'line':   {'none': True},
                                'points': points,
                                'data_labels': {'value': True, 'custom': custom_labels[i]},
                            })
                            start_row = 16
                        else:
                            start_row = 1
                        worksheet.insert_chart(start_row+15*i, len(df.columns)+col*8, chart, {'x_offset': 25, 'y_offset': 10})
                if len(cols) > 1:
                    for i, title in enumerate(['Magnitude [dB]', 'Phase [degree]']):
                        overlay_charts[i].set_x_axis({
                            'name': 'Frequency [kHz]',
                            'label_position': 'low',
                            'log_base': 10,
                            'display_units': 'thousands',
                            'display_units_visible': False,
                            'major_gridlines': {'visible': True},
                            'major_gridlines': {'visible': False},
                            'minor_tick_mark': 'none',
                            'min': df['Frequency'].min()
                        })
                        overlay_charts[i].set_y_axis({
                            'name': title,
                            'major_gridlines': {'visible': True}
                        })
                        overlay_charts[i].set_legend({'none': False, 'position': 'bottom'})
                        overlay_charts[i].set_title({'name': title})
                        worksheet.insert_chart(1, len(df.columns)+i*8, overlay_charts[i], {'x_offset': 25, 'y_offset': 10})

                writer.save()
                

def num2SIunit(num, string = True):
    SIunit = {
        'p': 1e-12,
        'n': 1e-9,
        'u': 1e-6,
        'm': 1e-3,
        '': 1,
        'k': 1e3,
        'M': 1e6,
        'G': 1e9,
        'T': 1e12
    }
    for unit_, exponent_ in SIunit.items():
        if abs(num) >= exponent_:
            unit, exponent = unit_, exponent_
        else:
            return '{:.2f}{}'.format(num/exponent, unit) if string else (num/exponent, unit)
    return str(num) if string else (num, '')


if __name__ == '__main__':
    main()