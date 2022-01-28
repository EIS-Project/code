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
        self.DUT = 10e3
        self.df = pd.read_excel(
            r'C:\Users\Public\Documents\Projects\EIS\Impedance_measurement\MSMT_data\parasitic inductance.xlsx')
        self.start_idx = self.df.loc[self.df['Z']
                                     < self.df.loc[:10, 'Z'].mean() * 0.9].index[0]
        self.end_idx = self.df.loc[self.df['Z']
                                   == self.df['Z'].min()].index[0] - 1

    @staticmethod
    def solve_L_equation():
        from sympy import symbols, solve, Eq, sqrt, pi
        L, f, X, m = symbols('L, f, X, m')
        print(solve(Eq(1 / (2 * pi * f * L)**m, X), L))
        print(solve(Eq((2 * pi * f * L)**m, X), L))
        #     print(
        #         solve(Eq(1 / (2 * pi * df.loc[i, 'freq'] * L)**(slope), df.loc[i, 'Z']), L))

    @property
    def C(self):
        wL = self.df.loc[self.start_idx:self.end_idx]
        log_wL = log(wL)
        m = abs(linregress(log_wL['freq'], log_wL['Z']).slope)
        val = 0
        for i in log_wL.index:
            f = self.df.loc[i, 'freq']
            X = self.df.loc[i, 'Z']
            val += (2**(-m) * pi**(-m) / X)**(1 / m) / f
        print(m)
        return val / log_wL.shape[0]

    @property
    def L(self):
        wC = self.df.loc[self.end_idx + 2:]
        log_wC = log(wC)
        m = abs(linregress(log_wC['freq'], log_wC['Z']).slope)
        val = 0
        for i in log_wC.index:
            f = self.df.loc[i, 'freq']
            X = self.df.loc[i, 'Z']
            val += (2**(-m) * pi**(-m) * X)**(1 / m) / f
        print(m)
        return val / log_wC.shape[0]

    def compensate(self):
        Esr = self.df['Z'].min()
        w = 2 * pi * self.df['freq']
        C = self.C
        L = self.L
        Z = self.df['Z']
        m = 1.77711436743729
        n = 3.70928587644783
        X = (1 / (w * C)**n - (w * L)**m)
        Z_final = Z - (np.sqrt(Esr ** 2 + X ** 2) -
                       np.sqrt(self.DUT ** 2 + X ** 2))
        return Z_final


if __name__ == '__main__':
    rlc = RLC()
    t1 = time.perf_counter()
    # print(rlc.compensate(), time.perf_counter() - t1)

    plt.loglog(rlc.df['freq'], rlc.compensate())
    plt.xlabel('freq')
    plt.ylabel('XL')
    plt.show()
