import os, sys
from notion_client import Client
from dotenv import load_dotenv

from data_collection import collect
from data_presentation import present

import pandas as pd

def main():

    if len(sys.argv) < 2:
        print("Usage: \n\t python main.py [notion_url]\n")
        return

    page_urls = (sys.argv[1:])
    page_ids = list(map(lambda url: url[-32:], page_urls))
    page_ids_formatted = []

    for page_id in page_ids:
        page_ids_formatted.append(collect.get_pageid_from_url(page_id[-32:]))

    load_dotenv()

    notion_secret = os.environ.get("SECRET")

    notion = Client(auth=notion_secret)

    cell_pos = (0, 0)
    table_objs = []
    filepath = "./tables_notion.xlsx"

    table_obj_dict = {}

    for page_id_formatted in page_ids_formatted:
        page_content = notion.blocks.children.list(block_id=str(page_id_formatted))
        table_obj = collect.get_table_objects(notion, page_content)

        table_objs += table_obj

    for table_obj in table_objs:
        table_obj_constructed = present.build_table_from_obj(table_obj)

        if table_obj_constructed[0][0] in table_obj_dict.keys():
            table_obj_dict[table_obj_constructed[0][0]] += table_obj_constructed[2:]
        else:
            table_obj_dict[table_obj_constructed[0][0]] = table_obj_constructed[1:]

        for obj_dict_index, obj_dict in enumerate(table_obj_dict[table_obj_constructed[0][0]][1:]):
            if obj_dict_index == 0:
                obj_dict[0] = '1st'
            elif obj_dict_index == 1:
                obj_dict[0] = '2nd'
            elif obj_dict_index == 2:
                obj_dict[0] = '3rd'
            else:
                obj_dict[0] = f'{obj_dict_index+1}th'

    with pd.ExcelWriter(filepath, 
                        engine="xlsxwriter", 
                        mode="w") as writer:

        cell_pos = (cell_pos[0], 0)

        workbook = writer.book
        chart_worksheet = None

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

                min_col_letter = present.number_to_alphabet(min_col)
                max_col_letter = present.number_to_alphabet(max_col)

                catmin_col_letter = present.number_to_alphabet(catmin_col)
                catmax_col_letter = present.number_to_alphabet(catmax_col)

                linechart.add_series({
                    'name': series_headers[serie_index+1],
                    'values': f'=Sheet!${min_col_letter}${min_row}:'+
                                      f'{max_col_letter}${max_row}',
                    'categories': f'=Sheet!${catmin_col_letter}${catmin_row}:'+
                                          f'{catmax_col_letter}${catmax_row}',
                })

            chart_worksheet.insert_chart('A1', linechart)

            cell_pos = (cell_pos[0], cell_pos[1] + len(values[0]))

if __name__ == '__main__':
    main()
