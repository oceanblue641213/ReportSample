from typing import List
from bson import ObjectId

def get_ghg_reports_data(table, company_id: str, fields: List[str]) -> dict:
    query = {'companyId': ObjectId(company_id)}
        
    projection = {field: 1 for field in fields}
    projection['_id'] = 1  # 始終包含_id
    
    return table.find_one(query, projection)