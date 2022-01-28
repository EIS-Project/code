from impedance.models.circuits import circuits
import pandas as pd
import numba
import time
from impedance import preprocessing
from impedance.models.circuits.circuits import CustomCircuit
from matplotlib import pyplot as plt
from math import pi
import numpy as np
from scipy.sparse import data
import os
from impedance import preprocessing
from impedance.models.circuits import CustomCircuit, fitting
from impedance.visualization import plot_nyquist, plot_bode, plot_altair, plot_residuals
from impedance.validation import linKK
import matplotlib


def impedance_pltot():
    start = 50
    stop = 4e6
    steps = 501
    step = (stop - start) // steps
    f = np.geomspace(start=1e3, stop=10e6, num=steps,
                     endpoint=True, dtype=None, axis=0)
    w = 2 * pi * f
    L = 3.3e-6
    C = 100e-9
    X = w * L + 1 / w / C
    R = 0
    z = np.sqrt(R**2 + X**2)
    print(z)
    plt.plot(f, z)
    plt.xscale('log')
    import pandas as pd
    df = pd.DataFrame({"Freq": f, 'Z': z})
    df.to_csv('simulation.csv', index=False)


def main():
    print(1 / (2 * pi * 38e3 * 21.46e3))
    print(1 / (pi * (52.57 - 9.81) * 1e3 * 38.1e3))  # 1/wR=C->R/2=1/2wC
    # 1/2*Rp = 1/wC


class Hybrid_RC:
    def __init__(self) -> None:
        self.Rs = {'1k': 500, '10k': 10e3, '75k': 74180}
        self.Rp = {'1k': 4.935e3, '10k': 42.76e3, '75k': 315677}
        self.fmax = {'1k': 38e3, '10k': 38e3, '75k': 43.651e3}

    def sweep_M(self, DUT):
        use_Lin_kk = True
        freq, Z = preprocessing.readCSV(
            f'./{DUT}.csv')
        # Z = np.conjugate(Z)
        freq, Z = preprocessing.ignoreBelowX(freq, Z)
        circuit = 'R0-p(R1, C1)'
        Cp = 1 / (pi * self.fmax[DUT] * self.Rp[DUT])
        #
        initial_guess = [self.Rs[DUT], self.Rp[DUT], Cp]  # [Cp, 30753*2]
        circuit = CustomCircuit(circuit.replace(
            ' ', ''), initial_guess=initial_guess)
        MvsMu = [[], []]
        for i in range(3):
            fig = plt.figure(figsize=(14, 7))
            gs = fig.add_gridspec(3, 3)
            # circuit.fit(freq, Z)
            for max_M in range(1, 4):
                M, mu, Z_linKK, res_real, res_imag = linKK(
                    freq, Z, c=None, max_M=max_M + i * 3, fit_type='complex', add_cap=False)
                print(f'{M}, {mu}')
                MvsMu[0].append(M)
                MvsMu[1].append(mu)
                Z_fit = Z_linKK  # circuit.predict(freq)
                print(
                    '\nCompleted Lin-KK Fit\nM = {:d}\nmu = {:.2f}'.format(M, mu))

                ax = fig.add_subplot(gs[:2, max_M - 1])

                ax2 = fig.add_subplot(gs[2, max_M - 1])

                plot_nyquist(ax, Z, fmt='o')

                plot_nyquist(ax, Z_fit, fmt='-')
                res_meas_real = (Z - Z_fit).real / np.abs(Z)
                res_meas_imag = (Z - Z_fit).imag / np.abs(Z)
                plot_residuals(ax2, freq, res_real,
                               res_imag, y_limits=(min([min(res_meas_imag), min(res_meas_real)]) * 100, max([max(res_meas_imag), max(res_meas_real)]) * 100))
                # circuit.plot(f_data=freq, Z_data=Z_fit, kind='bode')
                # circuit.plot(f_data=freq, Z_data=Z, kind='bode')
                ax.set_title(f'{M=}')
                ax.legend(['Data', 'Fit'])

                def mkfunc(x, pos): return '%1.1fM' % (
                    x * 1e-6) if x >= 1e6 else '%1dK' % (x * 1e-3) if x >= 1e3 else '%1d' % x
                mkformatter = matplotlib.ticker.FuncFormatter(mkfunc)
                ax.xaxis.set_major_formatter(mkformatter)
                ax.yaxis.set_major_formatter(mkformatter)
                ax2.set_title(f'µ={mu:.2f}')
                ax2.legend(['ΔRe', 'ΔIm'])

            plt.tight_layout(pad=150)
            plt.subplots_adjust(left=None, bottom=None, right=None,
                                top=None, wspace=.5, hspace=0.01)
            plt.savefig(os.path.join(
                r'C:\Users\Public\Documents\Projects\EIS\captures', f'-M={i*3+1}-{i*3+3}'))
        fig = plt.figure(figsize=(14, 7))
        # left, bottom, width, height (range 0 to 1)
        axes = fig.add_axes([0.1, 0.1, 0.8, 0.8])
        axes.plot(MvsMu[0], MvsMu[1], 'r')
        axes.set_xlabel('M')
        axes.set_ylabel('µ')
        axes.set_title('µ vs M')
        for xy in zip(MvsMu[0], MvsMu[1]):                                       # <--
            axes.annotate('(%s, %.2f)' % xy, xy=xy,
                          textcoords='data')  # <--

        axes.grid()
        plt.savefig(os.path.join(
            r'C:\Users\Public\Documents\Projects\EIS\captures', 'mu-vs-M'))
        print(fitting.rmse(Z, Z_fit))
        df = pd.DataFrame({'freq': freq, 'Z': np.abs(Z), 'Z.real': Z.real,
                           'Z.imag': Z.imag, 'Z_fit': np.abs(Z_fit), 'Z_fit.real': Z_fit.real, 'Z_fit.imag': Z_fit.imag, 'res_meas_real': res_meas_real, 'res_meas_imag': res_meas_imag, 'KK real': Z_linKK.real, 'KK imag': Z_linKK.imag, 'KK res_real': res_real, 'KK res_imag': res_imag})
        df['initial guess'] = f"""{circuit.circuit}, {circuit.initial_guess}"""
        df.to_csv(os.path.join(r'C:\Users\Public\Documents\Projects\EIS\Impedance_measurement\Nyquist Plots',
                  f'{DUT}_Nyquist Plot.csv'), index=False)

    @staticmethod
    def tap_water_lin_kk():
        # Load data from the example EIS result
        f, Z = preprocessing.readCSV('./tap-water.csv')

        # keep only the impedance data in the first quandrant
        f, Z = preprocessing.ignoreBelowX(f, Z)

        # mask = f < 1000
        # f = f[mask]
        # Z = Z[mask]
        M, mu, Z_linKK, res_real, res_imag = linKK(
            f, Z, c=.5, max_M=100, fit_type='complex', add_cap=True)
        print('\nCompleted Lin-KK Fit\nM = {:d}\nmu = {:.2f}'.format(M, mu))

        fig = plt.figure(figsize=(5, 8))
        gs = fig.add_gridspec(1, 1)
        ax1 = fig.add_subplot(gs[:, :])
        fig = plt.figure(figsize=(8, 3))
        gs = fig.add_gridspec(1, 1)
        ax2 = fig.add_subplot(gs[:, :])

        # plot original data
        plot_nyquist(ax1, Z, fmt='s')
        # plot measurement model
        plot_nyquist(ax1, Z_linKK, fmt='-', units='\Omega')

        ax1.legend(['Data', 'Lin-KK model'], loc=2, fontsize=12)
        # format axis
        ax1 = matplotlib.pyplot.gca()
        mkfunc = lambda x, pos: '%1.1fM' % (
            x * 1e-6) if x >= 1e6 else '%1.1fK' % (x * 1e-3)
        mkformatter = matplotlib.ticker.FuncFormatter(mkfunc)
        ax1.yaxis.set_major_formatter(mkformatter)

        # Plot residuals
        plot_residuals(ax2, f, res_real, res_imag, y_limits=(min([min(res_real), min(
            res_real)]) * 100, max([max(res_imag), max(res_imag)]) * 100))

        plt.tight_layout()
        plt.show()


Hybrid_RC1 = Hybrid_RC()
Hybrid_RC1.tap_water_lin_kk()

# def Real(self):
#     return self.Rs+self.Rp/(1+(self.w*self.Rp*self.Cp)**2)

# def X(self):
#     return -self.w*self.Rp**2*self.Cp/(1+(self.w*self.Rp*self.Cp)**2)
