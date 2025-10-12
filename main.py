import os, csv
from notion_client import Client
from dotenv import load_dotenv

def build_row(num_columns, cells, row=[]): 
    for index, cell in enumerate(cells):

        try:
            cell_content = cell[0]['text']['content']
        except IndexError:
            cell_content = ''

        row.insert(len(row), cell_content)

        if index >= num_columns:
            continue

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

        with open(f"data{table_obj_index}.csv", 'w', newline='') as f:
            writer = csv.writer(f)

            table_contents = notion.blocks.children.list(block_id=table_id)
            
            for table_content in table_contents['results']:

                try:
                    cells = table_content['table_row']['cells']
                except KeyError:
                    cells = ''

                num_columns = len(cells)

                row = build_row(num_columns, cells)

                writer.writerow(row)

                row.clear()

        print(f"Wrote data{table_obj_index}.csv file")


if __name__ == '__main__':
    main()
