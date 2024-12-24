from typing import List
from bson import ObjectId

def get_company_assets_data(table, company_ids: List[str]) -> List[dict]:
    query = {
        'company': {'$in': [ObjectId(id) for id in company_ids]}
    }
    
    return list(table.find(query))