import os, sys
from notion_client import Client
from dotenv import load_dotenv

from data_collection import collect
from data_presentation import present

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

    for table_obj in table_objs: 
        table = present.build_table_from_obj(table_obj)
        title = table[0][0]

        workbook = present.create_linechart_from_data(table, title)
        workbook.close()

if __name__ == '__main__':
    main()
