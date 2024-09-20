from openpyxl import load_workbook

def process_excel_file(file_path):
    # Load workbook in read-write mode
    wb = load_workbook(file_path)
    ws = wb['inter_106320750']

    area_cor_column = 'E'  # E
    a_int_os_column = 'N'  # N
    geocodigo_column = 'F'  # F
    area_int_column = 'M'  # M
    enq_legal_column = 'L'  # L

    distinct_values = {}

    # Row-wise iteration with direct access to cells
    for row in ws.iter_rows(min_row=2, values_only=False):
        cell_value = row[ws[geocodigo_column][0].column - 1].value
        row_num = row[0].row
        area_int_value = row[ws[area_int_column][0].column - 1].value
        enq_legal_value = row[ws[enq_legal_column][0].column - 1].value
        area_cor_value = row[ws[area_cor_column][0].column - 1].value
        a_int_os_value = row[ws[a_int_os_column][0].column - 1].value

        if cell_value not in distinct_values:
            distinct_values[cell_value] = {
                'sum_1': area_int_value if area_int_value is not None else 0,
                'single_value': enq_legal_value if enq_legal_value is not None else 0,
                'difference_left': 0,
                'rows': [row_num]
            }
        else:
            distinct_values[cell_value]['sum_1'] += area_int_value if area_int_value is not None else 0
            distinct_values[cell_value]['rows'].append(row_num)

    # Distribute differences and update the worksheet
    for value, data in distinct_values.items():
        data['difference'] = round(data['single_value'] - data['sum_1'], 4)

        remaining_difference = data['difference']
        for row_num in data['rows']:
            a_int_os_value = ws[f'{a_int_os_column}{row_num}'].value
            area_cor_value = ws[f'{area_cor_column}{row_num}'].value

            a_int_os_value = a_int_os_value if a_int_os_value is not None else 0
            max_increase = round(area_cor_value - a_int_os_value, 4)

            addition = min(remaining_difference, max_increase)
            ws[f'{a_int_os_column}{row_num}'].value = round(a_int_os_value + addition, 4)

            remaining_difference -= addition
            if remaining_difference <= 0:
                break

        data['difference_left'] = remaining_difference

    wb.save('file_updated.xlsx')
