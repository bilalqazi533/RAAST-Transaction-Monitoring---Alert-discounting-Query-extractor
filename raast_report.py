import pandas as pd
import re

# -----------------------------
# Step 1: Load the CSV file
# -----------------------------
input_file_path = 'fds_pmts_SADAPKKA_262106080600_262206080600.csv'  # Update if needed
data = pd.read_csv(input_file_path)

# -----------------------------
# Count Declined transactions
# -----------------------------
declined_count = data[data['STATUS'] == 'Declined'].shape[0]

# -----------------------------
# Step 2: Extract date and time from the input file name
# -----------------------------
input_file_name = input_file_path.split('/')[-1]
match = re.search(
    r"_(\d{2})(\d{2})(\d{4})(\d{4})_(\d{2})(\d{2})(\d{4})(\d{4})\.csv$",
    input_file_name
)

if match:
    start_day, start_month, start_year, start_time = match.group(1), match.group(2), match.group(3), match.group(4)
    end_day, end_month, end_year, end_time = match.group(5), match.group(6), match.group(7), match.group(8)
    output_file_name = (
        f"RAAST_REPORT_{start_day}_{start_month}_{start_year}_{start_time}_"
        f"TO_{end_day}_{end_month}_{end_year}_{end_time}.xlsx"
    )
else:
    raise ValueError(f"Input file name format is incorrect: {input_file_name}")

# -----------------------------
# Step 3: Filter the data
# -----------------------------
incoming_data = data[data['CREDITOR PARTICIPANT'] == 'SADAPKKA']
outgoing_data = data[data['SENDER PARTICIPANT'] == 'SADAPKKA']

incoming_filtered = incoming_data[incoming_data['STATUS'] != 'Declined']
outgoing_filtered = outgoing_data[outgoing_data['STATUS'] != 'Declined']

# -----------------------------
# Step 4: Incoming Pivot Table
# -----------------------------
pivot_incoming_data = incoming_filtered.pivot_table(
    index=['CREDITOR ACCOUNT'],
    values=['AMOUNT', 'RISK WEIGHT'],
    aggfunc={'AMOUNT': 'sum', 'RISK WEIGHT': 'max'}
).reset_index()

count_incoming_data = (
    incoming_filtered.groupby('CREDITOR ACCOUNT')
    .size()
    .reset_index(name='COUNTA of CREDITOR ACCOUNT')
)

pivot_incoming_data = pd.merge(
    pivot_incoming_data,
    count_incoming_data,
    on='CREDITOR ACCOUNT'
)

pivot_incoming_data['Comments'] = pivot_incoming_data.apply(
    lambda row: 'QTBS'
    if (row['AMOUNT'] >= 10000 and row['COUNTA of CREDITOR ACCOUNT'] >= 30)
    or (row['RISK WEIGHT'] == 200 and row['COUNTA of CREDITOR ACCOUNT'] > 5)
    or (row['AMOUNT'] >= 30000 and row['RISK WEIGHT'] == 200)
    else '',
    axis=1
)

pivot_incoming_data = pivot_incoming_data.sort_values(
    by='COUNTA of CREDITOR ACCOUNT',
    ascending=False
)

# -----------------------------
# Step 5: Outgoing Pivot Table
# -----------------------------
pivot_outgoing_data = outgoing_filtered.pivot_table(
    index=['DEBTOR ACCOUNT'],
    values=['AMOUNT', 'RISK WEIGHT'],
    aggfunc={'AMOUNT': 'sum', 'RISK WEIGHT': 'max'}
).reset_index()

count_outgoing_data = (
    outgoing_filtered.groupby('DEBTOR ACCOUNT')
    .size()
    .reset_index(name='COUNTA of DEBTOR ACCOUNT')
)

pivot_outgoing_data = pd.merge(
    pivot_outgoing_data,
    count_outgoing_data,
    on='DEBTOR ACCOUNT'
)

pivot_outgoing_data['Comments'] = pivot_outgoing_data.apply(
    lambda row: 'QTBS'
    if (row['AMOUNT'] > 10000 and row['COUNTA of DEBTOR ACCOUNT'] >= 30)
    or (row['RISK WEIGHT'] == 200 and row['COUNTA of DEBTOR ACCOUNT'] > 5)
    or (row['AMOUNT'] >= 30000 and row['RISK WEIGHT'] == 200)
    else '',
    axis=1
)

pivot_outgoing_data = pivot_outgoing_data.sort_values(
    by='COUNTA of DEBTOR ACCOUNT',
    ascending=False
)

# -----------------------------
# Step 6: Save full report Excel
# -----------------------------
with pd.ExcelWriter(output_file_name, engine='openpyxl') as writer:
    data.to_excel(writer, sheet_name='Original', index=False)
    incoming_data.to_excel(writer, sheet_name='Incoming', index=False)
    outgoing_data.to_excel(writer, sheet_name='Outgoing', index=False)
    pivot_incoming_data.to_excel(writer, sheet_name='Incoming Pivot', index=False)
    pivot_outgoing_data.to_excel(writer, sheet_name='Outgoing Pivot', index=False)

print(f"Excel file saved: {output_file_name}")

# -----------------------------
# Step 7: SUS, TOTAL, FP calculations
# -----------------------------
total_references = data['REFERENCE'].count()

sus_incoming = pivot_incoming_data[
    pivot_incoming_data['Comments'] == 'QTBS'
]['COUNTA of CREDITOR ACCOUNT'].sum()

sus_outgoing = pivot_outgoing_data[
    pivot_outgoing_data['Comments'] == 'QTBS'
]['COUNTA of DEBTOR ACCOUNT'].sum()

total_sus = sus_incoming + sus_outgoing
fp = total_references - total_sus

print(f"Incoming SUS: {sus_incoming}")
print(f"Outgoing SUS: {sus_outgoing}")
print(f"Total SUS: {total_sus}")
print(f"TOTAL: {total_references}")
print(f"FP: {fp}")
print(f"Declined Transactions: {declined_count}")

# -----------------------------
# Step 8: QTBS Filtering and Description
# -----------------------------
qtbs_incoming_data = pivot_incoming_data[pivot_incoming_data['Comments'] == 'QTBS'].copy()
qtbs_outgoing_data = pivot_outgoing_data[pivot_outgoing_data['Comments'] == 'QTBS'].copy()

qtbs_incoming_data["Customer's Name"] = qtbs_incoming_data['CREDITOR ACCOUNT'].map(
    dict(zip(incoming_filtered['CREDITOR ACCOUNT'], incoming_filtered['CREDITOR NAME']))
)

qtbs_outgoing_data["Customer's Name"] = qtbs_outgoing_data['DEBTOR ACCOUNT'].map(
    dict(zip(outgoing_filtered['DEBTOR ACCOUNT'], outgoing_filtered['DEBTOR NAME']))
)

qtbs_incoming_data.rename(columns={
    'CREDITOR ACCOUNT': 'PAN',
    'AMOUNT': 'Billing amount',
    'COUNTA of CREDITOR ACCOUNT': 'Number of Transactions'
}, inplace=True)

qtbs_outgoing_data.rename(columns={
    'DEBTOR ACCOUNT': 'PAN',
    'AMOUNT': 'Billing amount',
    'COUNTA of DEBTOR ACCOUNT': 'Number of Transactions'
}, inplace=True)

transaction_date_span = f"{start_day}/{start_month}/{start_year}"
qtbs_incoming_data['Transaction Date span (start)'] = transaction_date_span
qtbs_outgoing_data['Transaction Date span (start)'] = transaction_date_span

columns_order = [
    'PAN',
    "Customer's Name",
    'Number of Transactions',
    'Billing amount',
    'Transaction Date span (start)'
]

qtbs_incoming_data = qtbs_incoming_data[columns_order]
qtbs_outgoing_data = qtbs_outgoing_data[columns_order]

fixed_date = "21/06/2026"

qtbs_incoming_data['Description'] = [
    f'=B{i+2} & " received a sum of PKR " & D{i+2} & " from multiple account holders in " & C{i+2} & " transactions on the {fixed_date}. Please find out the relationship between the user and these account holders and the purpose of these transactions."'
    for i in range(len(qtbs_incoming_data))
]

qtbs_outgoing_data['Description'] = [
    f'=B{i+2} & " sent a sum of PKR " & D{i+2} & " to multiple account holders in " & C{i+2} & " transactions on the {fixed_date}. Please find out the relationship between the user and these account holders and the purpose of these transactions."'
    for i in range(len(qtbs_outgoing_data))
]

# -----------------------------
# Step 9: Save QTBS Excel
# -----------------------------
filtered_output_file_name = (
    f"QTBS_REPORT_{start_day}_{start_month}_{start_year}_{start_time}_"
    f"TO_{end_day}_{end_month}_{end_year}_{end_time}.xlsx"
)

with pd.ExcelWriter(filtered_output_file_name, engine='openpyxl') as writer:
    qtbs_incoming_data.to_excel(writer, sheet_name='Incoming QTBS', index=False)
    qtbs_outgoing_data.to_excel(writer, sheet_name='Outgoing QTBS', index=False)

print(f"Filtered QTBS file saved: {filtered_output_file_name}")