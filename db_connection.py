from typing import Dict
from pymongo import MongoClient
from pymongo.collection import Collection

class MongoDBConnection:
    def __init__(self, connection_string: str, database_name: str):
        self._client = MongoClient(connection_string)
        self._db = self._client[database_name]
        
    @property
    def site_modules(self) -> Collection:
        return self._db['site_modules']
    
    @property
    def company_assets(self) -> Collection:
        return self._db['company_assets']
    
    @property
    def onboarding_ratings(self) -> Collection:
        return self._db['onbording-ratings']
    
    @property
    def total_Emission(self) -> Collection:
        return self._db['total_Emission']
    
    @property
    def translations(self) -> Collection:
        return self._db['translations']
    
    def get_all_collections(self) -> Dict[str, Collection]:
        return {
            'site_modules': self.site_modules,
            'company_assets': self.company_assets,
            'onboarding_ratings': self.onboarding_ratings,
            'total_Emission': self.total_Emission,
            'translations': self.translations
        }
    
    def close(self):
        if hasattr(self, '_client'):
            self._client.close()
            
    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

# 使用範例
def create_connection(connection_string: str = 'your_connection_string') -> MongoDBConnection:
    return MongoDBConnection(connection_string, 'db_prod')