import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
import math

# Read the Excel file
tr = pd.read_excel('tr.xlsx')
df = pd.read_excel('regs.xlsx')

reg_C = df.iloc[4:256, 2].tolist()  # Column C (index 2), rows 90-257
dereg_C = df.iloc[4:256, 8].tolist()  # Column I (index 8), rows 90-257



# yearly increase in businesses 
surplus = [reg_C[i] - dereg_C[i] for i in range(0,257-5) ]

jr = []
for i in range(math.floor(len(surplus)/12)):
    jr.append(sum(surplus[i:i+13]))

surplus = jr
print(surplus)

# Extract the columns
trk_C = tr.iloc[5:26, 2].tolist()  # Column C (index 2), rows 90-257
trk_D = tr.iloc[5:26, 1].tolist()  # Column C (index 2), rows 90-257

trk_C = [y/x for x,y in zip(trk_D,trk_C)]

print(trk_C)


# Calculate correlation coefficient
correlation = pd.Series(trk_C).corr(pd.Series(surplus))

plt.figure(figsize=(10, 6))
plt.scatter(trk_C, surplus, c='#1E90FF', s=80, alpha=0.7, edgecolors='white')

# Add a trend line
slope, intercept, r_value, p_value, std_err = stats.linregress(trk_C, surplus)
line = slope * pd.Series(trk_C) + intercept
plt.plot(trk_C, line, color='#FF6347', lw=2, linestyle='--', alpha=0.8)

# Customize the plot
plt.title('TRK vs Business Surplus', fontsize=16, fontweight='bold')
plt.xlabel('TRK', fontsize=12)
plt.ylabel('Business Surplus', fontsize=12)
plt.grid(True, linestyle=':', alpha=0.7)

# Add a text box with R-squared value
r_squared = r_value**2
plt.text(0.05, 0.95, f'R² = {r_squared:.3f}', transform=plt.gca().transAxes, 
         fontsize=10, verticalalignment='top')

# Add a text box with correlation coefficient
plt.text(0.05, 0.90, f'Correlation = {correlation:.3f}', transform=plt.gca().transAxes, 
         fontsize=10, verticalalignment='top')

# Adjust layout and display the plot
plt.tight_layout()
plt.show()
