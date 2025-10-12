import os, csv
from notion_client import Client
from dotenv import load_dotenv

def build_table(table_objs):

    table = []
    table_row = []

    for table_content in table_objs:

        try:
            cells = table_content['table_row']['cells']
        except KeyError:
            cells = ''

        table_row = build_row(cells)
        table.append(table_row)

    return table

def build_row(cells): 

    row = []

    for cell in cells:

        try:
            cell_content = cell[0]['text']['content']
        except IndexError:
            cell_content = ''

        row.insert(len(row), cell_content)

    return row

def main():

    load_dotenv()

    notion_secret = os.environ.get("SECRET")
    page_id = os.environ.get("PAGE_ID")

    notion = Client(auth=notion_secret)

    page_content = notion.blocks.children.list(block_id=str(page_id))

    table_objects = list(filter(lambda obj: obj['type'] == 'table', page_content['results']))
    table_ids = list(map(lambda tab: tab['id'], table_objects))

    for table_obj_index, table_id in enumerate(table_ids):

        filename = f"notion_table_{table_obj_index}.csv"

        with open(filename, 'w', newline='') as f:

            writer = csv.writer(f)

            table_contents = notion.blocks.children.list(block_id=table_id)

            table = build_table(table_contents['results'])

            writer.writerows(table)

        print(f"Wrote file: {filename}")


if __name__ == '__main__':
    main()
