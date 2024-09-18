from openpyxl import load_workbook

wb = load_workbook('INTER_106320750.xlsx')
ws = wb['inter_106320750']

distinct_column = 'F'
sum_column_1 = 'M'
single_value_column = 'L'  # Column to take the value from one row (not sum)

# Dictionary to store distinct values, list of row numbers, the sum of column 1, the value of column C, and the subtraction result
distinct_values = {}

# Iterate through the rows in the distinct column (excluding header if any)
for row in ws.iter_rows(min_col=ws[distinct_column][0].column, max_col=ws[distinct_column][0].column, min_row=2, values_only=False):
    cell_value = row[0].value
    row_num = row[0].row
    
    # Skip null (None) values
    if cell_value is None:
        continue

    # Get the value from the first sum column (B) for the same row
    sum_value_1 = ws[f'{sum_column_1}{row_num}'].value
    # Get the value from column C for the same row
    single_value = ws[f'{single_value_column}{row_num}'].value

    # Handle cases where sum column values or the single value column are None (null)
    sum_value_1 = sum_value_1 if sum_value_1 is not None else 0
    single_value = single_value if single_value is not None else 0

    # If the distinct value already exists, append the row and continue summing column B
    if cell_value in distinct_values:
        distinct_values[cell_value]['rows'].append(row_num)
        distinct_values[cell_value]['sum_1'] += sum_value_1
        # Only take the value of column C from the first occurrence (don't overwrite it)
    else:
        # Initialize the entry with the row number, sum from column B, and the single value from column C
        distinct_values[cell_value] = {
            'rows': [row_num],
            'sum_1': sum_value_1,  # Sum of values from column B
            'single_value': single_value  # Take value from column C (only from one row)
        }

# After collecting data, calculate the difference for each distinct value
for value, data in distinct_values.items():
    # Subtract sum_1 from single_value (b - a) and format to 4 decimal places
    data['difference'] = round(data['single_value'] - data['sum_1'], 4)

# Print the distinct values, their rows, sum of column B, the single value from column C, and the subtraction result
for value, data in distinct_values.items():
    print(f"Geocodigo: {value}, Enq. Legal: {data['single_value']}, Sum: {data['sum_1']}, Difference: {data['difference']}, Rows: {data['rows']}")