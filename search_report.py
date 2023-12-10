import pandas as pd

# Load the master list from masterlist.txt
with open('masterlist.txt', 'r') as file:
    master_list = [line.strip().lower() for line in file]

# Load the Excel report
excel_file = 'ODBC_export-20231130.xlsx'
df = pd.read_excel(excel_file)

# Handle cases where 'Endpoint' is not a string (e.g., float)
df['Host'] = df['Endpoint'].apply(lambda endpoint: str(endpoint).split('.')[0].lower() if isinstance(endpoint, str) else endpoint)

# Create a new column 'HostExists' indicating whether the host exists in the master list
df['HostExists'] = df['Host'].isin(master_list)

# Filter the rows where the host exists in the master list
result_df = df[df['HostExists']]

# Drop the 'Host' and 'HostExists' columns before saving to the results file
result_df = result_df.drop(['Host', 'HostExists'], axis=1)

# Write the result to results.txt
result_file = 'ODBCVersions.csv'
result_df.to_csv(result_file, index=False)

print(f"Results written to {result_file}")
