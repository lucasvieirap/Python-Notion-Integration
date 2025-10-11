import os
from notion_client import Client
from dotenv import load_dotenv

def main():

    load_dotenv()

    notion_secret = os.environ.get("SECRET")
    block_id = os.environ.get("BLOCK_ID")

    notion = Client(auth=notion_secret)

    results = notion.blocks.children.list(block_id=str(block_id))

    for result in results['results']:

        num_columns = len(result['table_row']['cells'])

        for index, cell in enumerate(result['table_row']['cells']):

            if cell[0]['type'] == "text":

                if index >= num_columns - 1:
                    print(cell[0]['text']['content'])

                elif index <= num_columns - 2:
                    print(cell[0]['text']['content'], end=" | ")

if __name__ == '__main__':
    main()
