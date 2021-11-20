from tkinter import filedialog as fd
import os
import pandas as pd
from pathlib import Path


def rearrage_data():

    Path(os.path.join(os.path.dirname(__file__), 'AD2_formatted')).mkdir(parents=True, exist_ok=True)

    filetypes = (
            ('excel csv', '*.csv'),
            # ('all files', '')
        )

    files = fd.askopenfilenames(title='Open a file',
            initialdir=os.path.dirname(__file__),filetypes=filetypes)

    for file in files:
        try:
            filename = os.path.basename(file)
            # print(filename)
            if pd.read_csv(file, nrows=1).columns[0] in ['#Digilent WaveForms Network Analyzer - Bode']:   
                nrows = 18
                df = rearrange(file, nrows)
            elif pd.read_csv(file, nrows=1).columns[0] == '#Digilent WaveForms Impedance Analyzer':
                nrows = 28
                df = rearrange(file, nrows)
            else:
                df = pd.read_csv(file)
            df.to_csv(os.path.join(os.path.dirname(__file__), 'AD2_formatted', filename), encoding='utf-8', index=False)
        except:
            print(filename, 'failed')

def rearrange(file, nrows):    
    info = pd.read_csv(file, nrows=nrows)
    data = pd.read_csv(file, skiprows=nrows+1)
    df = pd.concat([data, info], axis=1)
    return df

if __name__ == '__main__':
    rearrage_data()
