from typing import Dict, Any, List, Optional
import redis
from pymongo import MongoClient
from datetime import datetime
import json

class I18nService:
    def __init__(
        self,
        mongodb_uri: str,
        db_name: str = 'RD-Nxmap',  # 這是資料庫名稱
        collection_name: str = 'new_translations',  # 這是集合名稱
        redis_host: str = 'localhost',
        redis_port: int = 6379,
        redis_password: str = '12345678',
        default_lang: str = 'en'
    ):
        # MongoDB 設定
        self.mongo_client = MongoClient(mongodb_uri)
        self.db = self.mongo_client[db_name]  # 根據 db_name 來選擇資料庫
        self.collection = self.db[collection_name]  # 根據 collection_name 來選擇集合
        
        # Redis 設定
        self.redis_client = redis.Redis(
            host=redis_host,
            port=redis_port,
            password=redis_password,
            decode_responses=True
        )
        
        self.default_lang = default_lang
        self.cache_ttl = 3600  # Redis 快取時效，單位為秒

    def _get_translation_key(self, lang: str, key: str) -> str:
        """生成 Redis 快取的鍵值"""
        return f"i18n:{lang}:{key}"

    def _load_translations_to_redis(self, lang: str):
        """從 MongoDB 載入特定語言的所有翻譯到 Redis"""
        for aaa in self.collection.find({}):
            print(aaa)
            
        translations = self.collection.find({"lang": lang})
        
        pipeline = self.redis_client.pipeline()
        for trans in translations:
            redis_key = self._get_translation_key(lang, trans['key'])
            pipeline.set(redis_key, trans['value'], ex=self.cache_ttl)
        pipeline.execute()

    def get_translation(self, key: str, lang: str) -> str:
        """
        獲取翻譯，優先從 Redis 查詢，如果沒有則從 MongoDB 查詢並更新到 Redis
        """
        redis_key = self._get_translation_key(lang, key)
        
        # 嘗試從 Redis 獲取翻譯
        cached_value = self.redis_client.get(redis_key)
        if cached_value:
            return cached_value
        
        # 查詢條件，過濾出 lang = 'tw' 的文檔
        query = {'lang': lang}

        # 查詢指定文檔
        doc = self.collection.find_one(query)
        if doc:
            # 提取 translations 中對應 key 的值
            translation_value = doc.get('translations', {}).get(key)
            
            # 將翻譯存入 Redis
            self.redis_client.set(
                redis_key,
                translation_value,
                ex=self.cache_ttl
            )
            return translation_value
        else:
            print(f"Document for language '{lang}' not found")
        
        # 如果找不到翻譯，返回原始 key
        return key

    def translate_document(self, doc: Dict[str, Any], lang: str) -> Dict[str, Any]:
        """
        翻譯整個文件
        """
        result = doc.copy()
        
        for key, value in doc.items():
            if isinstance(value, dict):
                result[key] = self.translate_document(value, lang)
            elif isinstance(value, list):
                result[key] = [
                    self.translate_document(item, lang) if isinstance(item, dict)
                    else item
                    for item in value
                ]
            elif isinstance(value, str):
                result[key] = self.get_translation(key, lang)
        #  and key.endswith('_i18n')
        return result

    def translate_documents(self, docs: List[Dict[str, Any]], lang: str) -> List[Dict[str, Any]]:
        """
        批量翻譯多個文件
        """
        return [self.translate_document(doc, lang) for doc in docs]

    def refresh_translations(self, lang: str):
        """
        重新載入指定語言的所有翻譯到 Redis
        """
        # 清除該語言的所有快取
        pattern = f"i18n:{lang}:*"
        keys = self.redis_client.keys(pattern)
        if keys:
            self.redis_client.delete(*keys)
        
        # 重新載入
        self._load_translations_to_redis(lang)

    def update_translation(self, key: str, lang: str, value: str):
        """
        更新翻譯（同時更新 MongoDB 和 Redis）
        """
        # 更新 MongoDB
        self.db.translations.update_one(
            {"key": key, "lang": lang},
            {
                "$set": {
                    "value": value,
                    "updated_at": datetime.utcnow()
                }
            },
            upsert=True
        )
        
        # 更新 Redis
        redis_key = self._get_translation_key(lang, key)
        self.redis_client.set(redis_key, value, ex=self.cache_ttl)