from openpyxl.chart import (
    LineChart,
    Reference,
)

def number_to_alphabet(num: int) -> str:
    if num == 0:
        return ""

    return number_to_alphabet((num-1)//26) + chr((num-1)%26+ord("A"))

def create_chart_from_data(data_worksheet, chart_worksheet, chart_title, min_pos, max_pos, categories_min_pos, categories_max_pos, chart_pos):

    linechart = LineChart()

    linechart.title = chart_title
    linechart.style = 10
    linechart.x_axis.title = "Sets"

    (min_col, min_row) = min_pos
    (max_col, max_row) = max_pos

    (categories_min_col, categories_min_row) = categories_min_pos
    (categories_max_col, categories_max_row) = categories_max_pos

    data = Reference(data_worksheet, 
                     min_col=min_col, 
                     min_row=min_row,
                     max_col=max_col,
                     max_row=max_row)

    categories = Reference(data_worksheet, 
                     min_col=categories_min_col, 
                     min_row=categories_min_row,
                     max_col=categories_max_col,
                     max_row=categories_max_row)

    linechart.add_data(data, titles_from_data=False)
    linechart.set_categories(categories)

    column = number_to_alphabet(chart_pos[1]+1)
    row = chart_pos[0]+1

    chart_worksheet.add_chart(linechart, f'{column}{row}')

    return linechart

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
