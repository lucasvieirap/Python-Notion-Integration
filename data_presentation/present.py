import pandas as pd

def number_to_alphabet(num: int) -> str:
    if num == 0:
        return ""

    return number_to_alphabet((num-1)//26) + chr((num-1)%26+ord("A"))

def create_chart(writer, table_obj_dict):
    cell_pos = (0, 0)

    workbook = writer.book

    for name, values in table_obj_dict.items():
        table_data = [[name], *values]
        table_df = pd.DataFrame(data=table_data)
        table_df.to_excel(writer, 
                          startrow=cell_pos[0], 
                          startcol=cell_pos[1],
                          index=False,
                          sheet_name="Sheet")

        values_list = []
        series_list = []
        for row in values[1:]:
            values_list += row[1:]

        series_headers = values[0]
        series_count = len(values[0])-1

        for i in range(1, len(values[0])):
            series_list.append(list(map(lambda x: None if i >= len(x) else x[i], values[1:])))

        max_value = max(list(filter(lambda x: x != None, values_list)))

        chart_sheet_name = f"{name} CHART"

        chart_worksheet = workbook.add_worksheet(chart_sheet_name)

        linechart = workbook.add_chart({ 'type': 'line' })
        linechart.set_title({ 'name': name })
        linechart.set_style(13)
        linechart.set_y_axis({
            'name': 'Reps',
            'min': 0,
            'max': max_value+(max_value/3),
        })

        linechart.set_size({
            'x_scale': 2.3,
            'y_scale': 2.3,
        })

        for serie_index in range(0, series_count):

            (min_col, min_row) = (cell_pos[1] + (serie_index + 2), cell_pos[0]+4)
            (max_col, max_row) = (cell_pos[1] + (serie_index + 2), cell_pos[0]+1 + table_df.shape[0])

            (catmin_col, catmin_row) = (cell_pos[1]+1, 
                                        cell_pos[0]+4)
            (catmax_col, catmax_row) = (
                cell_pos[1]+1, 
                cell_pos[0]+1 + table_df.shape[0]
            )

            min_col_letter = number_to_alphabet(min_col)
            max_col_letter = number_to_alphabet(max_col)

            catmin_col_letter = number_to_alphabet(catmin_col)
            catmax_col_letter = number_to_alphabet(catmax_col)

            linechart.add_series({
                'name': series_headers[serie_index+1],
                'values': f'=Sheet!${min_col_letter}${min_row}:'+
                                  f'{max_col_letter}${max_row}',
                'categories': f'=Sheet!${catmin_col_letter}${catmin_row}:'+
                                      f'{catmax_col_letter}${catmax_row}',
            })

        chart_worksheet.insert_chart('A1', linechart)

        cell_pos = (cell_pos[0], cell_pos[1] + len(values[0]))

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
