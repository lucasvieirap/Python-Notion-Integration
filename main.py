import os, csv
from notion_client import Client
from dotenv import load_dotenv

def main():

    load_dotenv()

    notion_secret = os.environ.get("SECRET")
    page_id = os.environ.get("PAGE_ID")

    notion = Client(auth=notion_secret)

    page_content = notion.blocks.children.list(block_id=str(page_id))

    table_objects = list(filter(lambda obj: obj['type'] == 'table', page_content['results']))

    for table_obj_index, obj in enumerate(table_objects):

        with open(f"data{table_obj_index}.csv", 'w', newline='') as f:
            writer = csv.writer(f)

            table_id = obj['id']
            table_content = notion.blocks.children.list(block_id=table_id)
            
            for field in table_content['results']:

                try:
                    cells = field['table_row']['cells']
                except KeyError:
                    cells = ''

                num_columns = len(cells)

                row = []
                for index, cell in enumerate(cells):

                    try:
                        cell_content = cell[0]['text']['content']
                    except IndexError:
                        cell_content = ' '

                    row.insert(len(row), cell_content)
                    if index >= num_columns - 1:
                        writer.writerow(row)
                        row.clear()
                    if index <= num_columns - 2:
                        pass

        print(f"Wrote data{table_obj_index}.csv file")




if __name__ == '__main__':
    main()
