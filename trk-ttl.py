import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import math
import sys
from matplotlib.ticker import FuncFormatter

# Read the Excel file
df = pd.read_excel("regs.xlsx")

# Extract data
reg_C = df.iloc[4:264, 2].tolist()
dereg_C = df.iloc[4:266, 8].tolist() 
time_c = df.iloc[4:266, 0].dropna().tolist()  # Adjusted to match the range of reg_C and dereg_C

regs = []
deregs = []

for i in range(math.ceil(len(reg_C)/12)):
    regs.append(sum(reg_C[i:i+13]))

for i in range(math.ceil(len(dereg_C)/12)):
    deregs.append(sum(dereg_C[i:i+13]))

lower = [(regs[i] - deregs[i])*0.22 for i in range(len(regs))]
upper = [(regs[i] - deregs[i])*0.32 for i in range(len(regs))]
l_turkish = [(regs[i] - deregs[i])*0.15*0.22 for i in range(len(regs))]
u_turkish = [(regs[i] - deregs[i])*0.15*0.32 for i in range(len(regs))]

# Set up the plot style
sns.set_style("darkgrid")
sns.set_context("notebook")
sns.set_palette("deep")

# Create the figure and axis objects
fig, ax = plt.subplots(figsize=(12, 6))

# Plot the data
ax.plot(time_c, upper, label='Upper bound by migrants', linewidth=2, marker='o', markersize=4, color="#b89639")
ax.plot(time_c, lower, label='Lower bound by migrants', linewidth=2, marker='s', markersize=4, color="#a34a3e")
ax.plot(time_c, u_turkish, label='Lower bound by turkish migrants', linewidth=2, marker='s', markersize=4)
ax.plot(time_c, l_turkish, label='Upper bound by turkish migrants', linewidth=2, marker='o', markersize=4)

# Customize the plot
ax.set_xlabel('Year', fontsize=12, fontweight='bold')
ax.set_ylabel('Amount', fontsize=12, fontweight='bold')
ax.set_title('Surplus of Business Registrations - Total Migrants vs Turkish Migrants', fontsize=16, fontweight='bold')
ax.fill_between(time_c[:len(upper)], lower, upper, alpha=0.3, color='#a34a3e')
ax.fill_between(time_c[:len(upper)], u_turkish, l_turkish, alpha=0.3, color='darkorange')

# Add legend
ax.legend(fontsize=10, loc='best')

# Rotate x-axis labels for better readability
plt.xticks(rotation=45, ha='right')

# Format y-axis ticks to show actual numbers
def format_func(value, tick_number):
    return f'{value:,.0f}'

ax.yaxis.set_major_formatter(FuncFormatter(format_func))

# Set y-axis ticks with 5000 spacing
y_min = min(min(lower), min(l_turkish))
y_max = max(max(upper), max(u_turkish))
y_ticks = range(int(y_min - (y_min % 10000)), int(y_max + 10000), 10000)
plt.yticks(y_ticks)

# Show the plot
plt.tight_layout()
plt.show()
