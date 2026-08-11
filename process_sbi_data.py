import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# ==========================================
# PART 1: DATA CLEANING & PII REMOVAL
# ==========================================

# 1. Load the raw data
df = pd.read_excel("DATA ANALYSIS.xlsx", sheet_name='real data')

# 2. Remove PII (Personally Identifiable Information)
if 'Name :' in df.columns:
    df = df.drop(columns=['Name :'])

# 3. Clean column names for professional presentation
clean_columns = []
for col in df.columns:
    # Remove the repetitive questionnaire formatting from columns
    if "For each statement" in col:
        clean_col = col.split('[')[-1].replace(']', '').strip()
        clean_columns.append(clean_col)
    else:
        clean_columns.append(col.strip())
df.columns = clean_columns

# 4. Save the sanitized dataset
df.to_excel('SBI_Retail_Investor_Research_Dataset.xlsx', index=False)
print("Sanitized dataset saved as 'SBI_Retail_Investor_Research_Dataset.xlsx'")

# ==========================================
# PART 2: CHART GENERATION
# ==========================================
sns.set_theme(style="whitegrid")

# --- Chart 1: Asset Preference ---
# Fix the comma issue within the Insurance category before splitting
assets_cleaned = df['What are your main investment options? (Select all that apply)'].astype(str).str.replace('Insurance (Life, Health,Term , or Other)', 'Insurance', regex=False)
asset_list = [item.strip() for sublist in assets_cleaned.str.split(',') if isinstance(sublist, list) for item in sublist]

# Count and filter out blanks/errors (like '99' or 'nan')
asset_counts = pd.Series(asset_list).value_counts().drop(['99', 'nan'], errors='ignore')

plt.figure(figsize=(10, 6))
sns.barplot(x=asset_counts.values, y=asset_counts.index, hue=asset_counts.index, palette="Blues_r", legend=False)
plt.title('Retail Investor Asset Preferences (Multiple Selections Allowed)')
plt.xlabel('Number of Investors')
plt.ylabel('Asset Class')
plt.tight_layout()
plt.savefig('chart1_asset_preference.png', dpi=300)
plt.close()
print("Chart 1 saved.")

# --- Chart 2: Income Distribution ---
income_counts = df['Monthly income (INR):'].value_counts()

plt.figure(figsize=(8, 5))
sns.barplot(x=income_counts.index, y=income_counts.values, hue=income_counts.index, palette="viridis", legend=False)
plt.title('Monthly Income Distribution of Respondents')
plt.xlabel('Income Bracket (INR)')
plt.ylabel('Count')
plt.xticks(rotation=15)
plt.tight_layout()
plt.savefig('chart2_income_distribution.png', dpi=300)
plt.close()
print("Chart 2 saved.")

# --- Chart 3: Chi-Square Summary ---
# Visualizing the Intention-Action gap from your hypothesis tests
chi_data = {
    'Hypothesis Test': [
        'Strategy vs. Satisfaction (Salaried)', 
        'Strategy vs. Satisfaction (Non-Salaried)', 
        'Strategy vs. Satisfaction (Combined)', 
        'Strategic Belief vs. Strategy Possession', 
        'Risk Awareness vs. Balanced Portfolio'
    ],
    'P-Value': [0.276, 0.200, 0.418, 0.673, 0.065]
}
df_chi = pd.DataFrame(chi_data)

plt.figure(figsize=(10, 6))
sns.barplot(x='P-Value', y='Hypothesis Test', data=df_chi, hue='Hypothesis Test', palette="dark:b", legend=False)
# Add the 0.05 significance threshold line
plt.axvline(x=0.05, color='red', linestyle='--', label='Significance Threshold (0.05)')
plt.title('Chi-Square Hypothesis Testing: Intention-Action Gap')
plt.xlabel('P-Value (Values > 0.05 Fail to Reject Null Hypothesis)')
plt.ylabel('')
plt.legend()
plt.tight_layout()
plt.savefig('chart3_chi_square_summary.png', dpi=300)
plt.close()
print("Chart 3 saved.")