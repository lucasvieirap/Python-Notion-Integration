import os, sys, csv
import pandas as pd
from notion_client import Client
from dotenv import load_dotenv

def build_table(table_objs, title):

    table = []
    table_row = []
    table_title = []

    for table_content in table_objs:

        try:
            cells = table_content['table_row']['cells']
        except KeyError:
            cells = ''

        table_row = build_row(cells)
        table.append(table_row)

    table_title.append(title)
    for _ in range(0, (len(table[0])-1)):
        table_title.append('')

    table.insert(0, table_title)

    return table

def build_row(cells): 

    row = []

    for cell in cells:

        try:
            cell_content = cell[0]['text']['content']
        except IndexError:
            cell_content = ''

        row.append(cell_content)

    return row

def export_to_xlsx(filespath, target):
    with pd.ExcelWriter(target) as writer:
        for filename in os.listdir(filespath):
            if filename.endswith('.csv'):
                filepath = os.path.join(filespath, filename)
                df = pd.read_csv(filepath)
                sheet_name = os.path.splitext(filename)[0]
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Created {target} successfully")

def main():

    load_dotenv()

    if len(sys.argv) < 2:
        print("Usage: \n\t python main.py [notion_url]\n")
        return

    char_list = list(sys.argv[1][-32:])
    char_list.insert(8, '-')
    char_list.insert(13, '-')
    char_list.insert(18, '-')
    char_list.insert(23, '-')
    page_id = "".join(char_list)

    notion_secret = os.environ.get("SECRET")

    notion = Client(auth=notion_secret)

    page_content = notion.blocks.children.list(block_id=str(page_id))

    table_titles = list(filter(lambda obj: obj['type'] == 'paragraph' and obj['paragraph']['color'] == 'gray_background', page_content['results']))

    table_objects = list(filter(lambda obj: obj['type'] == 'table', page_content['results']))
    table_ids = list(map(lambda tab: tab['id'], table_objects))

    for table_obj_index, table_id in enumerate(table_ids):

        table_contents = notion.blocks.children.list(block_id=table_id)
        title = table_titles[table_obj_index]['paragraph']['rich_text'][0]['plain_text']
        filename = f"notion_table_{title}.csv"

        with open(filename, 'w', newline='') as f:

            writer = csv.writer(f)

            table = build_table(table_contents['results'], title)

            writer.writerows(table)

        print(f"Wrote file: {filename}")

    export_to_xlsx("./", "./notion_tables_sheets.xlsx")


if __name__ == '__main__':
    main()
