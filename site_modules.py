from typing import List
from bson import ObjectId

def get_site_modules_data(table, target_id: str) -> List[dict]:
    query = {
        '$or': [
            {'_id': ObjectId(target_id)},
            {'parentCompanyId': ObjectId(target_id)}
        ]
        }
        
    projection = {
        '_id': 1,
        'companyName': 1,
        'levelStatus': 1
    }
    
    return list(table.find(query, projection))