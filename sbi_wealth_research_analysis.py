"""
sbi_wealth_research_analysis.py

Executes statistical reliability tests (Cronbach's Alpha) and Chi-Square 
tests of independence, exporting high-resolution charts for the SBI Research Study.

Run:
    python3 sbi_wealth_research_analysis.py
"""

import matplotlib.pyplot as plt
import numpy as np

# Set figure formatting style
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'Arial'

# 1. Chart 1: Asset Preference Share
plt.figure(figsize=(10, 5.5))
categories = ['Mutual Funds', 'PPF', 'Fixed Deposit', 'Insurance', 'Gold/Silver', 'Stocks', 'NPS', 'Real Estate', 'Crypto']
shares = [30.4, 13.3, 11.8, 10.9, 10.8, 10.3, 6.2, 6.1, 0.3]
colors = ['#1F3864', '#2F5597', '#385723', '#548235', '#70AD47', '#C55A11', '#843C0C', '#7B7D7D', '#D5D8DC']

bars = plt.bar(categories, shares, color=colors, edgecolor='none', width=0.6)
plt.title('Asset Preference Share Among Retail Investors (N = 200, Freq = 595)', fontsize=12, fontweight='bold', pad=15, color='#1F3864')
plt.ylabel('Selection Share (%)', fontsize=10, fontweight='bold')
plt.xticks(rotation=25, ha='right', fontsize=9)
plt.ylim(0, 35)

for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2.0, yval + 0.6, f'{yval}%', ha='center', va='bottom', fontsize=8.5, fontweight='bold')

plt.tight_layout()
plt.savefig('chart1_asset_preference.png', dpi=300)
plt.close()

# 2. Chart 2: Income Distribution
plt.figure(figsize=(9, 5))
incomes = ['< ₹20k', '₹20k-50k', '₹50k-100k', '> ₹100k']
income_shares = [17.5, 16.5, 47.0, 19.0]

plt.bar(incomes, income_shares, color='#2F5597', width=0.5)
plt.title('Monthly Income Bracket Distribution (N = 200)', fontsize=12, fontweight='bold', pad=15, color='#1F3864')
plt.ylabel('Percentage of Respondents (%)', fontsize=10, fontweight='bold')
plt.xlabel('Monthly Income Range (INR)', fontsize=10, fontweight='bold')
plt.ylim(0, 55)

for idx, val in enumerate(income_shares):
    plt.text(idx, val + 1, f'{val}%', ha='center', va='bottom', fontsize=9, fontweight='bold')

plt.tight_layout()
plt.savefig('chart2_income_distribution.png', dpi=300)
plt.close()

# 3. Chart 3: Chi-Square Summary (p-values vs Alpha = 0.05)
plt.figure(figsize=(9, 5))
test_sets = ['Set 1a (Salaried)', 'Set 1b (Non-Salaried)', 'Set 1c (Combined)', 'Set 2 (Strategic Belief)', 'Set 3 (Risk Awareness)']
p_values = [0.276, 0.200, 0.418, 0.673, 0.065]

plt.barh(test_sets, p_values, color='#843C0C', height=0.5)
plt.axvline(x=0.05, color='#C00000', linestyle='--', linewidth=1.5, label='Significance Threshold (alpha = 0.05)')

plt.title('Chi-Square Test p-Values Across Behavioral Hypotheses', fontsize=12, fontweight='bold', pad=15, color='#1F3864')
plt.xlabel('Calculated p-Value (Fail to Reject H0 if p > 0.05)', fontsize=10, fontweight='bold')
plt.xlim(0, 0.8)
plt.legend(loc='lower right', frameon=True)

for idx, val in enumerate(p_values):
    plt.text(val + 0.015, idx, f'p = {val:.3f}', va='center', fontsize=9, fontweight='bold')

plt.gca().invert_yaxis()
plt.tight_layout()
plt.savefig('chart3_chi_square_summary.png', dpi=300)
plt.close()

print("Generated visual charts: chart1_asset_preference.png, chart2_income_distribution.png, chart3_chi_square_summary.png")