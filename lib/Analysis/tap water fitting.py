from asyncio.proactor_events import constants
from operator import le
import matplotlib.pyplot as plt
import numpy as np
from impedance.visualization import plot_nyquist, plot_residuals
from impedance import preprocessing
from impedance.models.circuits import Randles, CustomCircuit, fitting
# R0-p(R1,C1)-Wo1
# randles = Randles(initial_guess=[80, 205, 2000e-12, 5000, 0.7])
# randlesCPE = Randles(
#     initial_guess=[80, 205, 70e-12, 70e-12, 1.03e+04, 6.32e-02], CPE=True)
# customCircuit = CustomCircuit(
#     initial_guess=[51, 212, 80e-12, 1.727e-3, .7], circuit='R_0-p(R_1,C_1)-CPE_1')
total_error = 1e3
# for i in np.arange(10e-9, 25e-9, 0.2e-9):
customCircuit = CustomCircuit(
    initial_guess=[73, 212, 17.4e-9, 1.028e6, 0.28, 1.727e-3, .7], circuit='R_0-p(R_1,C_1)-p(Wo_1,CPE_1)')
# customCircuit = CustomCircuit(
#     initial_guess=[49, 212, 18.5e-9], circuit='R_0-p(R_1,C_1)')
# customCircuit = CustomCircuit(
#     initial_guess=[280, 1.184e6, 0.79, 1.727e-3, .7], circuit='R_1-p(Ws_1,CPE_1)')
# customConstantCircuit = CustomCircuit(initial_guess=[None, .005, .1, .005, .1, .001, None],
#                                       constants={'R_0': 0.02, 'Wo_1_1': 200},
#                                       circuit='R_0-p(R_1,C_1)-p(R_2,C_2)-Wo_1')
print(customCircuit)
frequencies, Z = preprocessing.readCSV('./tap-water.csv')

# keep only the impedance data in the first quandrant
frequencies, Z = preprocessing.ignoreBelowX(frequencies, Z)
# randles.fit(frequencies, Z)
# randlesCPE.fit(frequencies, Z, global_opt=True)
customCircuit.fit(frequencies, Z)
# customConstantCircuit.fit(frequencies, Z)

# print(customCircuit)
f_pred = np.logspace(np.log10(max(frequencies)), np.log10(
    min(frequencies)), num=len(frequencies))

# randles_fit = randles.predict(frequencies)
# randlesCPE_fit = randlesCPE.predict(f_pred)
customCircuit_fit = customCircuit.predict(frequencies)
# customConstantCircuit_fit = customConstantCircuit.predict(f_pred)
res_real = (Z - customCircuit_fit).real / np.abs(Z)
res_imag = (Z - customCircuit_fit).imag / np.abs(Z)
# if abs(np.mean(res_real)) + abs(np.mean(res_imag)) < total_error:
#     total_error = abs(np.mean(res_real)) + abs(np.mean(res_imag))
#     best_Z = Z
#     alpha = i
# print(alpha)
fig, ax = plt.subplots(figsize=(8, 12))

plot_nyquist(ax, Z, fmt='.')
# plot_nyquist(ax, randles_fit, fmt='-')
# plot_nyquist(ax, randlesCPE_fit, fmt='-')
plot_nyquist(ax, customCircuit_fit, fmt='-')
# plot_nyquist(ax, customConstantCircuit_fit, fmt='-')

ax.legend(['Data', 'fit'])
# Plot residuals
fig, ax = plt.subplots(figsize=(7, 5))

plot_residuals(ax, frequencies, res_real, res_imag, y_limits=(min([min(res_real), min(
    res_imag)]) * 100, max([max(res_real), max(res_imag)]) * 100))
print(fitting.rmse(Z, customCircuit_fit))

# plt.show()
