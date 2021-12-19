from matplotlib import pyplot as plt
import math
import numpy as np
start = 1e3
stop = 10e6
steps = 501
step = (stop-start)//steps
f = np.geomspace(start=1e3, stop=10e6, num=steps, endpoint=True, dtype=None, axis=0)
w = 2*math.pi*f
L = 3.3e-6
C = 100e-9
X = w*L+1/w/C
R = 0
z = np.sqrt(R**2+X**2)
print(z)
plt.plot(f, z)
plt.xscale('log')
import pandas as pd
df = pd.DataFrame({"Freq":f, 'Z':z})
df.to_csv('simulation.csv', index=False)