from typing import List
from bson import ObjectId

def get_onbording_ratings_data(table, company_id: str) -> dict:
    query = {'companyId': company_id}
    projection = {
        '_id': 0,  # 排除 _id
        'companyId': 1,
        'type': 1,
        'baseYear': 1,
        'baseYearPeriodStart': 1,
        'baseYearPeriodEnd': 1,
        'rating': 1,
        'assurance': 1,
        'threshold': 1
    }
    
    return table.find_one(query, projection)