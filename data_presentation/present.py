import math
import pandas as pd

import xlsxwriter

def number_to_alphabet(num: int) -> str:
    if num == 0:
        return ""

    return number_to_alphabet((num-1)//26) + chr((num-1)%26+ord("A"))

def create_linechart_from_data(data, workbook, worksheet, pos) -> xlsxwriter.Workbook:

    alpha_column = number_to_alphabet(pos['col'] + 1)

    sheet = workbook.add_worksheet()

    chart = workbook.add_chart({ 'type': 'line' })
    chart.set_title({'name': f'=Sheet1!${ alpha_column }${ pos['row'] + 1 }'})

    number_format = workbook.add_format({'num_format': '#'})

    max = 0

    for row_index, row in enumerate(data):

        for cell_index, cell in enumerate(row): 

            cell_pos = { 'row': pos['row'] + row_index, 'col': pos['col'] + cell_index }

            try: 

                worksheet.write_number(cell_pos['row'], cell_pos['col'], float(cell), number_format)

                if max < float(cell):
                    max = math.floor(float(cell)+(float(cell)/3))

            except (TypeError, ValueError):

                worksheet.write(cell_pos['row'], cell_pos['col'], cell)

            if cell_index > 0 and row_index == 1: 

                current_column = number_to_alphabet(cell_pos['col'] + 1)
                last_row_index = len(data)

                chart.add_series({
                    'name': f'=Sheet1!${current_column}${pos['row']+2}',
                    'categories': f'=Sheet1!${alpha_column}${pos['row']+3}:${alpha_column}${pos['row']+last_row_index}',
                    'values': f'=Sheet1!${current_column}${pos['row']+3}:${current_column}${pos['row']+last_row_index}',
                    'marker': {'type': 'circle'}
                })

    chart.set_x_axis({
        'name': 'Set',
    })
    chart.set_y_axis({
        'name': 'Reps', 
        'min': 0,
        'max': max
    })

    chart.set_style(11)

    sheet.insert_chart('A1', chart)

    return workbook

def build_table_from_obj(table_obj) -> list:

    table_row_objs = table_obj['results']

    table = []
    table_row = []

    for row in table_row_objs:

        try:

            cells = row['table_row']['cells']

        except KeyError:

            cells = None

        table_row = build_row(cells)
        table.append(table_row)

    return table

def build_row(cells) -> list: 

    row = []

    for cell in cells:

        try:

            cell_content = cell[0]['text']['content']

        except IndexError:

            cell_content = None

        row.append(cell_content)

    return row
