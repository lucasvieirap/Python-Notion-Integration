import os, csv, string, math
import pandas as pd

import xlsxwriter
from xlsxwriter.color import Color

def create_csv_from_table_obj(data, filename):
    with open(filename, 'w', newline='') as f:

        writer = csv.writer(f)

        table = build_table_from_obj(data, filename.split('.')[0])
        writer.writerows(table)

    print(f"Wrote file: {filename}")

def create_linechart_from_data(data, filename) -> xlsxwriter.Workbook:

    workbook = xlsxwriter.Workbook(f"{filename}.xlsx")
    worksheet = workbook.add_worksheet()
    number_format = workbook.add_format({'num_format': '#'})

    chart = workbook.add_chart({'type': 'line'})
    chart.set_title({'name': '=Sheet1!$A$1'})

    max = 0

    for row_index, row in enumerate(data):
        for cell_index, cell in enumerate(row): 
            try: 
                worksheet.write_number(row_index, cell_index, float(cell), number_format)

                if max < float(cell):
                    max = math.floor(float(cell)+(float(cell)/3))

            except (TypeError, ValueError):
                worksheet.write(row_index, cell_index, cell)

            if cell_index > 0 and row_index == 1: 

                current_column = string.ascii_uppercase[cell_index]
                last_row_index = len(data)

                chart.add_series({
                    'name': f'=Sheet1!${current_column}$2',
                    'categories': f'=Sheet1!$A$3:$A${last_row_index}',
                    'values': f'=Sheet1!${current_column}$3:${current_column}${last_row_index}',
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

    worksheet.insert_chart('F2', chart)

    return workbook

def create_chart_from_csv(csvfilepath):

    csv_data = []

    with open(csvfilepath, newline='') as csvfile:

        reader = csv.reader(csvfile)

        for row in reader:
            csv_data.append(row)

    workbook = create_linechart_from_data(csv_data, 'data')
    workbook.close()

def build_table_from_obj(table_obj):

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

def build_row(cells): 

    row = []

    for cell in cells:

        try:
            cell_content = cell[0]['text']['content']
        except IndexError:
            cell_content = None

        row.append(cell_content)

    return row

def export_to_xlsx(filespath, target):
    with pd.ExcelWriter(target) as writer:
        for filename in os.listdir(filespath):
            if filename.endswith('.csv'):
                filepath = os.path.join(filespath, filename)
                df = pd.read_csv(filepath)
                sheet_name = os.path.splitext(filename)[0]
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Created {target} successfully")

