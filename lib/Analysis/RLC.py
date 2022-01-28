import pandas as pd
from sklearn.metrics import mean_squared_error
import numpy as np
from numpy import log10 as log
# from numpy import sqrt
from scipy.stats import linregress
from math import pi

from matplotlib import pyplot as plt
import time


def RMSE():
    df = pd.read_excel(
        r'C:\Users\Public\Documents\Projects\EIS\Impedance_measurement\MSMT_data\100k resistor compensation IA adapter.xlsx')
    rmse = {}
    for col in ['no compensation', 'open', 'close', 'open and close']:
        rmse[f'{col}'] = mean_squared_error(
            [100e3] * df[f'{col}'].shape[0], df[f'{col}'], squared=False)
        print((rmse['no compensation'] -
               rmse[f'{col}']) / rmse['no compensation'] * 100)


class RLC:
    def __init__(self) -> None:
        self.df = pd.read_excel(
            r'C:\Users\Public\Documents\Projects\EIS\Impedance_measurement\MSMT_data\parasitic inductance.xlsx')
        self.start_idx = self.df.loc[self.df['Xc']
                                     < self.df.loc[:10, 'Xc'].mean() * 0.9].index[0]
        self.end_idx = self.df.loc[self.df['Xc']
                                   == self.df['Xc'].min()].index[0] - 1

    @staticmethod
    def solve_L_equation():
        from sympy import symbols, solve, Eq, sqrt, pi
        L, f, X, m = symbols('L, f, X, m')
        print(solve(Eq(1 / (2 * pi * f * L)**m, X), L))
        print(solve(Eq((2 * pi * f * L)**m, X), L))
        #     print(
        #         solve(Eq(1 / (2 * pi * df.loc[i, 'freq'] * L)**(slope), df.loc[i, 'Xc']), L))

    @property
    def L(self):
        wL = self.df.loc[self.start_idx:self.end_idx]
        log_wL = log(wL)
        m = abs(linregress(log_wL['freq'], log_wL['Xc']).slope)
        val = 0
        for i in log_wL.index:
            f = self.df.loc[i, 'freq']
            X = self.df.loc[i, 'Xc']
            val += (2**(-m) * pi**(-m) / X)**(1 / m) / f
        print(m)
        return val / log_wL.shape[0]

    @property
    def C(self):
        wC = self.df.loc[self.end_idx + 2:]
        log_wC = log(wC)
        m = abs(linregress(log_wC['freq'], log_wC['Xc']).slope)
        val = 0
        for i in log_wC.index:
            f = self.df.loc[i, 'freq']
            X = self.df.loc[i, 'Xc']
            val += (2**(-m) * pi**(-m) * X)**(1 / m) / f
        print(m)
        return val / log_wC.shape[0]



rlc = RLC()
t1 = time.perf_counter()
print(rlc.C, rlc.L, time.perf_counter() - t1)


# plt.loglog(df['freq'], (1 / (2 * pi * df['freq'] * L)))
# plt.xlabel('freq')
# plt.ylabel('XL')
# plt.show()
