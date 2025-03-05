import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats

# Read the Excel file
df = pd.read_excel('regs.xlsx')
tr = pd.read_excel('tr.xlsx')

# Extract the columns
reg_C = df.iloc[88:256, 2].tolist()  # Column C (index 2), rows 90-257
dereg_C = df.iloc[89:256, 8].tolist()  # Column I (index 8), rows 90-257

#working turks
tr_wrk = tr.iloc[12:27, 2].tolist()

# yearly increase in businesses 
surplus = [reg_C[i] - dereg_C[i] for i in range(0,257-90) ]

print(reg_C)



