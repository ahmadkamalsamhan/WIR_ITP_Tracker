import pandas as pd

# ------------------------------
# 1. Load the CSV file
# ------------------------------
file_path = "ITP_activities_log.csv"  # Replace with your actual CSV file path
df = pd.read_csv(file_path)

# Clean column names
df.columns = df.columns.str.strip()

# ------------------------------
# 2. Group by ITP Reference
# ------------------------------
grouped = df.groupby('ITP Reference').agg(
    Submittal_Reference=('Submittal Reference', 'first'),
    Checklist_Reference=('Checklist Reference', 'first'),
    Activity_Nos=('Activity No.', lambda x: list(x)),
    Activity_Descriptions=('Activiy Description', lambda x: list(x))
).reset_index()

# ------------------------------
# 3. Optional: Combine Activity No. + Description in one string
# ------------------------------
grouped['Activities'] = grouped.apply(
    lambda row: [f"{no} - {desc}" for no, desc in zip(row['Activity_Nos'], row['Activity_Descriptions'])],
    axis=1
)

# Drop the separate columns if you just want the combined Activities column
grouped = grouped.drop(columns=['Activity_Nos', 'Activity_Descriptions'])

# ------------------------------
# 4. Save the grouped data to a new CSV
# ------------------------------
output_file = "ITP_grouped.csv"
grouped.to_csv(output_file, index=False)

print(f"Grouped ITP activities saved to {output_file}")
