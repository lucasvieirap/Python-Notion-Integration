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

    page_id = collect.get_pageid_from_url((sys.argv[1])[-32:])

    load_dotenv()

    notion_secret = os.environ.get("SECRET")

    notion = Client(auth=notion_secret)

    page_content = notion.blocks.children.list(block_id=str(page_id))

    table_objs = collect.get_table_objects(notion, page_content)

    pos = (0, 0)

    with pd.ExcelWriter("tables_notion.xlsx", engine="xlsxwriter", mode="w") as writer:

        for table_obj_index, table_obj in enumerate(table_objs): 

            table = present.build_table_from_obj(table_obj)

            sheet_name = f"{table[0][0]} CHART-{table_obj_index}"
            present.pandas_create_chart_from_data(writer, 
                                                  table, 
                                                  table[0][0], 
                                                  sheet_name, 
                                                  pos)

            pos = ( pos[0], pos[1] + len(table[0]) + 1)

if __name__ == '__main__':
    main()
