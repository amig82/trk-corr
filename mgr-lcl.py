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
total = [regs[i] - deregs[i] for i in range(len(regs))]

# Set up the plot style
sns.set_style("darkgrid")
sns.set_context("notebook")
sns.set_palette("deep")

# Create the figure and axis objects
fig, ax = plt.subplots(figsize=(12, 6))



# Set labels and title
ax.set_xlabel('Time')
ax.set_ylabel('Value')

# Plot the data
ax.plot(time_c, upper, label='Upper bound of registrations by migrants', linewidth=2, marker='o', markersize=4)
ax.plot(time_c, lower, label='Lower bound of registrations by migrants', linewidth=2, marker='s', markersize=4)
ax.plot(time_c, total, label='Total', linewidth=2, marker='s', markersize=4, color="#a34a3e")

# Customize the plot
ax.set_xlabel('Year', fontsize=12, fontweight='bold')
ax.set_ylabel('Amount', fontsize=12, fontweight='bold')
ax.set_title('Surplus of Business Registrations - Total vs Migrants', fontsize=16, fontweight='bold')
ax.fill_between(time_c[:len(upper)], lower, upper, alpha=0.3, color='darkorange')
# Add legend
ax.legend(fontsize=10, loc='best')

# Rotate x-axis labels for better readability
plt.xticks(rotation=45, ha='right')

# Format y-axis ticks to show actual numbers
def format_func(value, tick_number):
    return f'{value:,.0f}'

ax.yaxis.set_major_formatter(FuncFormatter(format_func))

# Show the plot
plt.tight_layout()
plt.show()
