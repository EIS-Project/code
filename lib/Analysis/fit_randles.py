from asyncio import constants
from impedance.visualization import plot_nyquist, plot_residuals
import matplotlib.pyplot as plt
from impedance.models.circuits import CustomCircuit
from impedance import preprocessing
import pandas as pd
import numpy as np
# Load data from the example EIS result
from heapq import nlargest, nsmallest
from math import pi
from time import perf_counter


class fit_randles:
    def __init__(self) -> None:
        # keep only the impedance data in the first quandrant
        frequencies, Z = self.read_excel()
        self.t1 = perf_counter()
        frequencies, Z = preprocessing.ignoreBelowX(frequencies, Z)
        self.frequencies, self.Z = frequencies[:-30], Z[:-30]

    @staticmethod
    def read_excel():
        df = pd.read_excel(
            r'C:\Users\Public\Documents\Projects\EIS\Impedance_measurement\MSMT_data\salt-water-ecm-fitting.xlsx', sheet_name='0.4ppm')
        df = df.iloc[:-20, :]
        f = df['Frequency (Hz)']
        Z = df['Trace Rs (Ω)'] + 1j * df['Trace Xs (Ω)']
        return np.array(f), np.array(Z)

    def fitting(self, frequencies, Z, circuit, initial_guess, fit=True):
        """fit individual Circuit elements

        Args:
            frequencies ([type]): [description]
            Z ([type]): [description]
            circuit ([type]): [description]
            initial_guess ([type]): [description]
            fit (bool, optional): [description]. Defaults to True.

        Returns:
            [type]: [description]
        """
        circuit = CustomCircuit(
            circuit, initial_guess=initial_guess)
        if fit:
            circuit.fit(frequencies, Z)
        Z_fit = circuit.predict(frequencies)
        print(circuit)
        fig, ax = plt.subplots()
        plot_nyquist(ax, Z, fmt='.', units='\Omega')
        plot_nyquist(ax, Z_fit, fmt='-', units='\Omega')
        plt.legend(['Data', 'Fit'])
        fig, ax = plt.subplots()
        res_meas_real = (Z - Z_fit).real / np.abs(Z)
        res_meas_imag = (Z - Z_fit).imag / np.abs(Z)
        plot_residuals(ax, frequencies, res_meas_real, res_meas_imag, y_limits=(min([min(res_meas_real), min(
            res_meas_imag)]) * 100, max([max(res_meas_real), max(res_meas_imag)]) * 100))
        return circuit.parameters_ if circuit.parameters_ is not None else circuit.initial_guess

    @property
    def parameters(self):
        """estimate Randle Circuit parameters from Nyquist Plot

        Returns:
            [tuple]: [parameters after fitting Randle Circuit]
        """
        Z = self.Z
        frequencies = self.frequencies
        der = np.diff(Z.imag) / np.diff(Z.real)
        idx = np.where(np.isin(der, nlargest(10, abs(der))))[0][0]
        Warburg_degree, Z_Warburg = np.polyfit(
            Z.real[:idx], Z.imag[:idx], deg=1)
        alpha = - np.degrees(np.arctan(Warburg_degree)) / 90
        w = 2 * pi * self.frequencies[np.where(Z.imag == max(Z.imag))[0][0]]
        # fit Warburg
        R1, CPE1, alpha = self.fitting(frequencies[:idx], Z[:idx], 'R1-CPE_1',
                                       [Z_Warburg, 1 / Z_Warburg, alpha])
        # fit semicircle
        C = 1 / (w * R1)
        R1, C = self.fitting(
            frequencies[idx:], Z[idx:], 'p(R1,C1)', [R1, C])
        # fit Randle's circuit
        R1, C, CPE1, alpha = self.fitting(
            frequencies, Z, 'p(R1,C1)-CPE1', [R1, C, CPE1, alpha], fit=True)
        print(f'{R1=}, {C=}, {CPE1=}, {alpha=}')
        print(f'process time: {perf_counter() - self.t1}s')
        plt.show()
        return (R1, C, CPE1, alpha)


if __name__ == '__main__':
    randles = fit_randles()
    randles.parameters
