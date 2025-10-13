import os
import pandas as pd

def build_table_from_obj(table_obj, title):

    table_row_objs = table_obj['results']

    table = []
    table_row = []
    table_title = []

    for row in table_row_objs:

        try:
            cells = row['table_row']['cells']
        except KeyError:
            cells = ''

        table_row = build_row(cells)
        table.append(table_row)

    table_title.append(title)

    for _ in range(0, (len(table[0])-1)):
        table_title.append('')

    table.insert(0, table_title)

    return table

def build_row(cells): 

    row = []

    for cell in cells:

        try:
            cell_content = cell[0]['text']['content']
        except IndexError:
            cell_content = ''

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

