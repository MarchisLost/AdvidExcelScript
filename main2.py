from openpyxl import load_workbook

# Load the workbook and select the sheet
wb = load_workbook('INTER_106320750.xlsx')
ws = wb['inter_106320750']

# Specify the necessary columns
Geocodigo = 'F'  # Distinct column
Area_Int = 'M'  # Column to sum (Area_Int)
Enq_Legal = 'L'  # Column to take a single value from (Enq_Legal)
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

    # Get the value from the first sum column (Area_Int) and the column Enq_Legal for the same row
    sum_value_1 = ws[f'{Area_Int}{row_num}'].value
    single_value = ws[f'{Enq_Legal}{row_num}'].value
    area_cor_value = ws[f'{area_cor_column}{row_num}'].value
    a_int_os_value = ws[f'{a_int_os_column}{row_num}'].value

    # Initialize values if they are None
    sum_value_1 = sum_value_1 if sum_value_1 is not None else 0
    single_value = single_value if single_value is not None else 0

    # If the distinct value already exists, append the row and continue summing column Area_Int
    if cell_value in distinct_values:
        distinct_values[cell_value]['rows'].append(row_num)
        distinct_values[cell_value]['sum_1'] += sum_value_1
    else:
        # Initialize the entry with the row number, sum from column Area_Int, and the single value from Enq_Legal
        distinct_values[cell_value] = {
            'rows': [row_num],
            'sum_1': sum_value_1,  # Sum of values from column Area_Int
            'single_value': single_value,  # Take value from Enq_Legal (only from one row)
            'old_A_INT_OS': [],  # Store old A_INT_OS values
            'new_A_INT_OS': [],  # Store new A_INT_OS values
            'difference_left': 0  # Initialize remaining difference
        }

    # Ensure that the old value of A_INT_OS is captured before the update
    distinct_values[cell_value]['old_A_INT_OS'].append(a_int_os_value)

# Calculate the difference for each distinct value
for value, data in distinct_values.items():
    data['difference'] = round(data['single_value'] - data['sum_1'], 4)  # Subtract (b - a)

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
        ws[f'{a_int_os_column}{row_num}'].value = new_value  # Update the cell

        # Store the new value of A_INT_OS in the dictionary
        distinct_values[value]['new_A_INT_OS'].append(new_value)

        # Update the remaining difference after the addition
        remaining_difference -= addition

        # If remaining_difference is 0, break the loop (no more rows need updating)
        if remaining_difference <= 0:
            break

    # Store any remaining difference in the dictionary
    data['difference_left'] = remaining_difference

# Save the modified workbook
wb.save('file_updated.xlsx')

# Print the distinct values, their rows, sum of column Area_Int, the single value from Enq_Legal, old and new A_INT_OS values, and the remaining difference
for value, data in distinct_values.items():
    print(f"Geocodigo: {value}, Enq. Legal: {data['single_value']}, Sum: {data['sum_1']}, Difference: {data['difference']}, Old A_INT_OS: {data['old_A_INT_OS']}, New A_INT_OS: {data['new_A_INT_OS']}, Difference left: {data['difference_left']:.4f}")
