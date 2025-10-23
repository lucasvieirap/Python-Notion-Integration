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

        present.create_chart(writer, table_obj_dict)

if __name__ == '__main__':
    main()
