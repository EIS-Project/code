import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog as fd

file = fd.askopenfilename(initialdir=os.path.dirname(__file__), title= "Please select a file:", filetypes= [('xlsx', '*.xlsx'),])
print(file)
Rref = 1e6
df = pd.read_excel(file)
Vp1 = np.power(10, df['V(p1) [dB]']/20)
Vp2 = np.power(10, df['V(p2) [dB]']/20)
df['Rdut'] = (Vp1-Vp2)/(Vp2/Rref)
# df.plot(kind='scatter', x='Frequency', y='Rdut', logx=True)
# plt.show()
writer = pd.ExcelWriter(os.path.join(r'C:\Users\User01\Desktop\EIS\Simulations', os.path.basename(file)))
df.to_excel(writer, sheet_name='impedance', index=False)
writer.save()