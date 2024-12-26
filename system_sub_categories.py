def get_system_sub_categories_data(table) -> dict:
    projection = {
        '_id': 1,
        'key': 1,
        'label': 1
    }
    
    result = [item for item in table.find({}, projection) if item.get('label')[0] != '1']
    
    return result