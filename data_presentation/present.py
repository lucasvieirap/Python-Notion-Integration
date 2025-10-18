import pandas as pd

def number_to_alphabet(num: int) -> str:
    if num == 0:
        return ""

    return number_to_alphabet((num-1)//26) + chr((num-1)%26+ord("A"))

def pandas_create_chart_from_data(writer, 
                                  data,
                                  chart_title,
                                  sheet_name,
                                  start_pos) -> None:

    table_df = pd.DataFrame(data=data)

    table_df.to_excel(writer, 
                      startrow=start_pos[0], 
                      startcol=start_pos[1],
                      index=False,
                      sheet_name="Sheet1")

    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)

    chart = workbook.add_chart({'type': 'line'})
    chart.set_title({'name': chart_title})

    max_value = max(list(map(lambda d: max(d[1:]), data[2:])))

    chart.set_x_axis({
        'name': 'Set',
    })
    chart.set_y_axis({
        'min': 0,
        'max': max_value + (max_value // 3)
    })

    (max_row, _) = table_df.shape

    serie = {'name': '', 'categories': [], 'values': []}

    for column_index, column in enumerate(data[1]):

        serie['name'] = column

        if column_index == 0:

            (category_pos_row, category_pos_col) = (start_pos[0] + 3, start_pos[1])
            (max_category_row, max_category_col) = (
                max_row, 
                category_pos_col
            )

            serie['categories'] = [
                'Sheet1',
                category_pos_row,
                category_pos_col,
                max_category_row,
                max_category_col,
            ]

            continue

        value_column_index = column_index - 1

        (value_pos_row, value_pos_col) = (
            start_pos[0] + 3, 
            start_pos[1] + 1 + value_column_index
        )
        (max_value_row, max_value_col) = (max_row, value_pos_col)

        serie['values'] = [
            "Sheet1", 
            value_pos_row, 
            value_pos_col, 
            max_value_row, 
            max_value_col,
        ]

        chart.add_series(serie)

    worksheet.insert_chart(0, 0, chart)

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

            cell_content = float(cell[0]['text']['content'])

        except (TypeError, ValueError):

            cell_content = cell[0]['text']['content']

        except (IndexError, KeyError):

            cell_content = None

        row.append(cell_content)

    return row
