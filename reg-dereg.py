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

# Set up the plot style
sns.set_style("darkgrid")
sns.set_context("notebook")
sns.set_palette("deep")

# Create the figure and axis objects
fig, ax = plt.subplots(figsize=(12, 6))

# Plot the data
ax.plot(time_c, regs, label='Business Registrations', linewidth=2, marker='o', markersize=4)
ax.plot(time_c, deregs, label='Business Deregistrations', linewidth=2, marker='s', markersize=4)

# Customize the plot
ax.set_xlabel('Year', fontsize=12, fontweight='bold')
ax.set_ylabel('Amount', fontsize=12, fontweight='bold')
ax.set_title('Registrations vs Deregistrations Over Time', fontsize=16, fontweight='bold')

# Add legend
ax.legend(fontsize=10, loc='best')

# Rotate x-axis labels for better readability
plt.xticks(rotation=45, ha='right')

# Format y-axis ticks to show actual numbers
def format_func(value, tick_number):
    return f'{value:,.0f}'

ax.yaxis.set_major_formatter(FuncFormatter(format_func))

# Tight layout to prevent clipping of labels
plt.tight_layout()

# Show the plot
plt.show()
