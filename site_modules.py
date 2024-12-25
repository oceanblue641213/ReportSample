from typing import List
from bson import ObjectId

def get_site_modules_data(table, target_id: str, fields: List[str]) -> List[dict]:
    """
    使用自定義欄位列表查詢資料
    
    Args:
        table: MongoDB collection對象
        target_id (str): 目標ID
        fields (List[str]): 要返回的欄位列表
        
    Returns:
        List[dict]: 符合條件的資料列表
    """
    
    query = {
        '$or': [
            {'_id': ObjectId(target_id)},
            {'parentCompanyId': ObjectId(target_id)}
        ]
        }
        
    projection = {field: 1 for field in fields}
    projection['_id'] = 1  # 始終包含_id
    
    return list(table.find(query, projection))