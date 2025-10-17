def get_pageid_from_url(urlstring):
    char_list = list(urlstring)
    char_list.insert(8, '-')
    char_list.insert(13, '-')
    char_list.insert(18, '-')
    char_list.insert(23, '-')
    page_id = "".join(char_list)
    return page_id

def get_table_objects(notion, dataset):

    list_objs = list(filter(lambda obj: obj['type'] == 'column_list' or obj['type'] == 'table', dataset['results']))
    list_ids = list(map(lambda tab: tab['id'], list_objs))

    data = []

    for index, obj in enumerate(list_objs):

        if obj['type'] == 'table':
            tab_obj = notion.blocks.children.list(block_id=str(obj['id']))
            data.append(tab_obj)
            continue

        column_objs = notion.blocks.children.list(block_id=str(list_ids[index]))

        for column in column_objs['results']:
            obj = notion.blocks.children.list(block_id=str(column['id']))
            if obj['results'][0]['type'] == 'table':
                tab_obj = notion.blocks.children.list(block_id=str(obj['results'][0]['id']))
                data.append(tab_obj)

    return data
