from openpyxl import load_workbook

wb = load_workbook('INTER_106320750.xlsx')
ws = wb['inter_106320750']

Geocodigo = 'F'
Area_Int = 'M'
Enq_Legal = 'L'  # Column to take the value from one row (not sum)

# # Dictionary to store distinct values, list of row numbers, the sum of column 1, the value of column C, and the subtraction result
# distinct_values = {}

# # Iterate through the rows in the distinct column (excluding header if any)
# for row in ws.iter_rows(min_col=ws[Geocodigo][0].column, max_col=ws[Geocodigo][0].column, min_row=2, values_only=False):
#     cell_value = row[0].value
#     row_num = row[0].row

#     # Skip null (None) values
#     if cell_value is None:
#         continue

#     # Get the value from the first sum column (B) for the same row
#     sum_value_1 = ws[f'{Area_Int}{row_num}'].value
#     # Get the value from column C for the same row
#     single_value = ws[f'{Enq_Legal}{row_num}'].value

#     # Handle cases where sum column values or the single value column are None (null)
#     sum_value_1 = sum_value_1 if sum_value_1 is not None else 0
#     single_value = single_value if single_value is not None else 0

#     # If the distinct value already exists, append the row and continue summing column B
#     if cell_value in distinct_values:
#         distinct_values[cell_value]['rows'].append(row_num)
#         distinct_values[cell_value]['sum_1'] += sum_value_1
#         # Only take the value of column C from the first occurrence (don't overwrite it)
#     else:
#         # Initialize the entry with the row number, sum from column B, and the single value from column C
#         distinct_values[cell_value] = {
#             'rows': [row_num],
#             'sum_1': sum_value_1,  # Sum of values from column B
#             'single_value': single_value  # Take value from column C (only from one row)
#         }

# # After collecting data, calculate the difference for each distinct value
# for value, data in distinct_values.items():
#     # Subtract sum_1 from single_value (b - a) and format to 4 decimal places
#     data['difference'] = round(data['single_value'] - data['sum_1'], 4)

# # Print the distinct values, their rows, sum of column B, the single value from column C, and the subtraction result
# for value, data in distinct_values.items():
#     print(f"Geocodigo: {value}, Enq. Legal: {data['single_value']}, Sum: {data['sum_1']}, Difference: {data['difference']}, Rows: {data['rows']}")

area_cor_column = 'E'  # Column AREA_COR
a_int_os_column = 'N'  # Column A_INT_OS

# Dictionary to store distinct values, rows, sum, difference, and difference_left
distinct_values = {}

# Iterate through the rows in the distinct column (excluding header if any)
for row in ws.iter_rows(min_col=ws[Geocodigo][0].column, max_col=ws[Geocodigo][0].column, min_row=2, values_only=False):
    cell_value = row[0].value
    row_num = row[0].row
    
    # Skip null (None) values
    if cell_value is None:
        continue

    # Get the value from the first sum column (B) and the column C for the same row
    sum_value_1 = ws[f'{Area_Int}{row_num}'].value
    single_value = ws[f'{Enq_Legal}{row_num}'].value
    area_cor_value = ws[f'{area_cor_column}{row_num}'].value
    a_int_os_value = ws[f'{a_int_os_column}{row_num}'].value

    # Handle cases where sum column values, single value column, area_cor, or a_int_os are None
    sum_value_1 = sum_value_1 if sum_value_1 is not None else 0
    single_value = single_value if single_value is not None else 0
    area_cor_value = area_cor_value if area_cor_value is not None else 0
    a_int_os_value = a_int_os_value if a_int_os_value is not None else 0

    # If the distinct value already exists, append the row and continue summing column B
    if cell_value in distinct_values:
        distinct_values[cell_value]['rows'].append(row_num)
        distinct_values[cell_value]['sum_1'] += sum_value_1
    else:
        # Initialize the entry with the row number, sum from column B, and the single value from column C
        distinct_values[cell_value] = {
            'rows': [row_num],
            'sum_1': sum_value_1,  # Sum of values from column B
            'single_value': single_value,  # Take value from column C (only from one row)
            'difference_left': 0  # Initialize remaining difference
        }

# Calculate the difference for each distinct value
for value, data in distinct_values.items():
    data['difference'] = round(data['single_value'] - data['sum_1'], 4)  # Subtract (b - a)

    # Distribute the difference to the rows
    remaining_difference = data['difference']
    for row_num in data['rows']:
        a_int_os_value = ws[f'{a_int_os_column}{row_num}'].value
        area_cor_value = ws[f'{area_cor_column}{row_num}'].value

        # If A_INT_OS is None, set it to 0
        a_int_os_value = a_int_os_value if a_int_os_value is not None else 0

        # Calculate the maximum amount that can be added to A_INT_OS without exceeding AREA_COR
        max_increase = round(area_cor_value - a_int_os_value, 4)

        # Add the smaller of remaining_difference or max_increase to A_INT_OS
        addition = min(remaining_difference, max_increase)
        ws[f'{a_int_os_column}{row_num}'].value = round(a_int_os_value + addition, 4)  # Update the cell

        # Update the remaining difference after the addition
        remaining_difference -= addition

        # If remaining_difference is 0, break the loop (no more rows need updating)
        if remaining_difference <= 0:
            break

    # Store any remaining difference in the dictionary
    data['difference_left'] = remaining_difference

# Save the modified workbook (if you want to save the changes to a new file)
wb.save('file_updated.xlsx')

# Print the distinct values, their rows, sum of column B, the single value from column C, and the remaining difference
for value, data in distinct_values.items():
    print(f"Geocodigo: {value}, Enq. Legal: {data['single_value']}, Sum: {data['sum_1']}, Difference: {data['difference']}, Rows: {data['rows']}, Difference left: {data['difference_left']:.4f}")
