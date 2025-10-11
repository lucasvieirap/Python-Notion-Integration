import os
from notion_client import Client
from dotenv import load_dotenv

def main():

    load_dotenv()

    notion_secret = os.environ.get("SECRET")
    page_id = os.environ.get("PAGE_ID")

    notion = Client(auth=notion_secret)

    page_content = notion.blocks.children.list(block_id=str(page_id))

    for object in page_content['results']:
        if object['type'] == 'table':
            table_id = object['id']
            table_content = notion.blocks.children.list(block_id=table_id)

            for table_results in table_content['results']:
                cells = table_results['table_row']['cells']
                columns = len(cells)
                for index, cell in enumerate(cells):
                    try:
                        cell_content = cell[0]['text']['content']
                    except IndexError:
                        cell_content = ' '
                    if index >= columns - 1:
                        print(cell_content)
                    if index <= columns - 2:
                        print(cell_content, end=" | ")

            print('=' * 25 + '\n')

if __name__ == '__main__':
    main()
