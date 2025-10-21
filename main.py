import os, sys
from notion_client import Client
from dotenv import load_dotenv

from data_collection import collect
from data_presentation import present

import pandas as pd
import openpyxl
from openpyxl.chart import (
    LineChart,
    Reference,
)
from openpyxl.chart.axis import Scaling

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
    chart_pos = (0, 0)

    table_objs = []

    filepath = "./tables_notion.xlsx"

    table_obj_dict = {}

    for page_id_formatted in page_ids_formatted:
        page_content = notion.blocks.children.list(block_id=str(page_id_formatted))
        table_obj = collect.get_table_objects(notion, page_content)

        table_objs += table_obj

    for table_obj in table_objs:
        table_obj_constructed = present.build_table_from_obj(table_obj)

        try:
            table_obj_dict[table_obj_constructed[0][0]] += table_obj_constructed[2:]
        except KeyError:
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

    # print(table_obj_dict)

    if not os.path.exists(filepath):
        workbook = openpyxl.Workbook()
        workbook.save(filepath)

    with pd.ExcelWriter(filepath, 
                        engine="openpyxl", 
                        if_sheet_exists="overlay",
                        mode="a") as writer:

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
            for value in values[1:]:
                values_list += value[1:]

            max_value = max(list(filter(lambda x: x != None, values_list)))

            linechart = LineChart()

            chart_sheet_name = f"{name} CHART"

            chart_worksheet = workbook.create_sheet(chart_sheet_name)

            (min_col, min_row) = (cell_pos[1] + 2, cell_pos[0]+3)
            (max_col, max_row) = (cell_pos[1] + table_df.shape[1], cell_pos[0]+1 + table_df.shape[0])

            (categories_min_col, categories_min_row) = (cell_pos[1]+1, 
                                                        cell_pos[0]+4)
            (categories_max_col, categories_max_row) = (
                cell_pos[1]+1, 
                cell_pos[0]+1 + table_df.shape[0]
            )

            linechart = LineChart()
            linechart.title = name
            linechart.style = 13
            linechart.y_axis.delete = False
            linechart.y_axis.title = "Reps"
            linechart.y_axis.scaling = Scaling(min=0, max=max_value+(max_value/3))

            data = Reference(workbook['Sheet'], 
                             min_col=min_col, 
                             min_row=min_row,
                             max_col=max_col,
                             max_row=max_row)

            categories = Reference(workbook['Sheet'], 
                                   min_col=categories_min_col, 
                                   min_row=categories_min_row,
                                   max_col=categories_max_col,
                                   max_row=categories_max_row)

            linechart.add_data(data, titles_from_data=True)
            linechart.set_categories(categories)

            chart_worksheet.add_chart(linechart, "A1")

            cell_pos = (cell_pos[0], cell_pos[1] + len(values[0]))

        workbook.save(filepath)

if __name__ == '__main__':
    main()
