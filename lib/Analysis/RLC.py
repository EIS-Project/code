"""
impedance Compensation for AD2 Impedance Analyzer measurement
"""
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
        r'C:\Users\Public\Documents\Projects\EIS\Impedance_measurement\MSMT_data\100k resistor compensation IA adapter.xlsx', sheet_name='Sheet2')
    rmse = {}
    for col in ['no compensation', 'open', 'close', 'open and close']:
        rmse[f'{col}'] = mean_squared_error(
            [100e3] * df[f'{col}'].shape[0], df[f'{col}'], squared=False)
        print((rmse['no compensation'] -
               rmse[f'{col}']) / rmse['no compensation'] * 100)


class Compensation:
    def __init__(self) -> None:
        self.DUT = 10e3
        self.df = pd.read_excel(
            r'C:\Users\Public\Documents\Projects\EIS\Impedance_measurement\MSMT_data\parasitic inductance.xlsx', sheet_name='Sheet1')
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
        w = self.df.loc[self.start_idx:self.end_idx]
        log_w = log(w)
        n = abs(linregress(log_w['freq'], log_w['Z']).slope)
        self.n = n
        val = (2 * pi)**(-1) / w['Z']**(1 / n) / w['freq']
        return np.mean(val.loc[abs(val - val.mean()) < val.std()])

    @property
    def L(self):
        w = self.df.loc[self.end_idx + 2:]
        log_w = log(w)
        m = abs(linregress(log_w['freq'], log_w['Z']).slope)
        self.m = m
        val = (2 * pi)**(-1) * w['Z']**(1 / m) / w['freq']
        return np.mean(val.loc[abs(val - val.mean()) < val.std()])

    @property
    def Z(self):
        Esr = self.df['Z'].min()
        w = 2 * pi * self.df['freq']
        C = self.C
        L = self.L

        Z = self.df['Z']
        m = self.m
        n = self.n
        print(f'{C=}, {L=}, {m=}, {n=}')
        X = (1 / (w * C)**n - (w * L)**m)
        Z_final = Z - (np.sqrt(Esr ** 2 + X ** 2) -
                       np.sqrt(self.DUT ** 2 + X ** 2))
        return Z_final

    @property
    def f(self):
        return self.df['freq']


if __name__ == '__main__':
    # comp = Compensation()
    # t1 = time.perf_counter()
    # # print(rlc.compensate(), time.perf_counter() - t1)

    # plt.semilogx(comp.f, comp.Z)
    # plt.xlabel('freq')
    # plt.ylabel('XL')
    # plt.show()
    print()
