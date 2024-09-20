from openpyxl import load_workbook

# Load the workbook and select the sheet
wb = load_workbook('INTER_106320750.xlsx')
ws = wb['inter_106320750']

# Specify the necessary columns
Geocodigo = 'F'  # Distinct column
Area_Int_column = 'M'  # Column to sum (Area_Int_column)
Enq_Legal_column = 'L'  # Column to take a single value from (Enq_Legal_column)
area_cor_column = 'E'  # Column AREA_COR
a_int_os_column = 'N'  # Column A_INT_OS

# Dictionary to store distinct values, rows, sum, difference, old and new values, and difference_left
distinct_values = {}

# Iterate through the rows in the distinct column (excluding header if any)
for row in ws.iter_rows(min_col=ws[Geocodigo][0].column, max_col=ws[Geocodigo][0].column, min_row=2, values_only=False):
    cell_value = row[0].value
    row_num = row[0].row

    # Skip null (None) values in the distinct column
    if cell_value is None:
        continue

    # Get the value from the first sum column (Area_Int_column) and the column Enq_Legal_column for the same row
    area_int = ws[f'{Area_Int_column}{row_num}'].value
    enq_legal = ws[f'{Enq_Legal_column}{row_num}'].value
    area_cor_value = ws[f'{area_cor_column}{row_num}'].value
    a_int_os_value = ws[f'{a_int_os_column}{row_num}'].value

    # Initialize values if they are None
    area_int = area_int if area_int is not None else 0
    enq_legal = enq_legal if enq_legal is not None else 0

    # If the distinct value already exists, append the row and continue summing column Area_Int_column
    if cell_value in distinct_values:
        distinct_values[cell_value]['rows'].append(row_num)
        distinct_values[cell_value]['sum_1'] += area_int
    else:
        # Initialize the entry with the row number, sum from column Area_Int_column, and the single value from Enq_Legal_column
        distinct_values[cell_value] = {
            'rows': [row_num],
            'sum_1': area_int,  # Sum of values from column Area_Int_column
            'enq_legal': enq_legal,  # Take value from Enq_Legal_column (only from one row)
            'old_A_INT_OS': [],  # Store old A_INT_OS values
            'new_A_INT_OS': [],  # Store new A_INT_OS values
            'difference_left': 0  # Initialize remaining difference
        }

    # Ensure that the old value of A_INT_OS is captured before the update
    distinct_values[cell_value]['old_A_INT_OS'].append(a_int_os_value)

# Calculate the difference for each distinct value
for value, data in distinct_values.items():
    data['difference'] = round(data['enq_legal'] - data['sum_1'], 4)  # Subtract (b - a)

    # Distribute the difference to the rows
    remaining_difference = data['difference']
    for row_num in data['rows']:
        a_int_os_value = ws[f'{a_int_os_column}{row_num}'].value
        area_cor_value = ws[f'{area_cor_column}{row_num}'].value

        # Calculate the maximum amount that can be added to A_INT_OS without exceeding AREA_COR
        max_increase = round(area_cor_value - a_int_os_value, 4)

        # Add the smaller of remaining_difference or max_increase to A_INT_OS
        addition = min(remaining_difference, max_increase)
        new_value = round(a_int_os_value + addition, 4)

        # If no addition is made, keep the original value
        if addition == 0:
            new_value = a_int_os_value

        # Update the cell value with the new value
        ws[f'{a_int_os_column}{row_num}'].value = new_value

        # Append the new or unchanged value to new_A_INT_OS list
        distinct_values[value]['new_A_INT_OS'].append(new_value)

        # Update the remaining difference after the addition
        remaining_difference -= addition

    # Store any remaining difference in the dictionary
    data['difference_left'] = remaining_difference

# Save the modified workbook
wb.save('file_updated.xlsx')

# Print the distinct values, their rows, sum of column Area_Int_column, the single value from Enq_Legal_column, old and new A_INT_OS values, and the remaining difference
for value, data in distinct_values.items():
    print(f"Geo: {value}, Enq.Legal: {data['enq_legal']}, Sum: {data['sum_1']}, Diff: {data['difference']}, Old_A_INT_OS: {data['old_A_INT_OS']}, New_A_INT_OS: {data['new_A_INT_OS']}, Diff left: {data['difference_left']:.4f}")
