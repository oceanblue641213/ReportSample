from typing import List
from bson import ObjectId

def get_new_total_emission_data(table, categories, company_ids: List[str]) -> List[dict]:
    projection = {
        '_id': 0,
        'report_data': 1
    }
    
    query = table.find({
        'companyId': {'$in': [id for id in company_ids]},
        "report_data.year": 2024
    },projection)
    
    result = [item for item in query]
    
    stationary_id = next(item['_id'] for item in categories if item['key'] == 'Stationary Combustion')
    mobile_id = next(item['_id'] for item in categories if item['key'] == 'Mobile Combustion')
    industrial_id = next(item['_id'] for item in categories if item['key'] == 'Industrial Processes')
    refrigerants_id = next(item['_id'] for item in categories if item['key'] == 'Refrigerants and Fugitives')
    
    # 找到所有 label 開頭為 2 的 _id
    categories2_ids = [item['_id'] for item in categories if item['label'].startswith('2')]
    # 找到所有 label 開頭為 3 的 _id
    categories3_ids = [item['_id'] for item in categories if item['label'].startswith('3')]
    # 找到所有 label 開頭為 4 的 _id
    categories4_ids = [item['_id'] for item in categories if item['label'].startswith('4')]
    
    # 初始化總和的字典
    totals = {
        'co2GWPEmissionValue': 0,
        'ch4GWPEmissionValue': 0,
        'n2oGWPEmissionValue': 0,
        'hfcsGWPEmissionValue': 0,
        'pfcsGWPEmissionValue': 0,
        'sf6GWPEmissionValue': 0,
        'nf3GWPEmissionValue': 0,
        'totalGWPEmissionValue': 0
    }
    
    ghg_emission_tables = {
        'totalStationaryCombustionValue': 0,
        'totalIndustrialProcesses': 0,
        'totalMobileCombustionValue': 0,
        'totalRefrigerantsFugitives': 0,
        'totalCategories2Value': 0,
        'totalCategories3Value': 0,
        'totalCategories4Value': 0
    }
    
    direct_ghg_emission = {
        'co2GWPEmissionValue': 0,
        'ch4GWPEmissionValue': 0,
        'n2oGWPEmissionValue': 0,
        'hfcsGWPEmissionValue': 0,
        'pfcsGWPEmissionValue': 0,
        'sf6GWPEmissionValue': 0,
        'nf3GWPEmissionValue': 0,
        'totalGWPEmissionValue': 0
    }
    
    # 遍歷每個文檔
    for doc in result:
        # 遍歷每個 report_data
        for report in doc['report_data']:
            # 取得 report_yearly_total 的值
            yearly_total = report['report_yearly_total']
            
            # 取得 report_yearly_direct 的值
            yearly_direct = report['report_yearly_direct']
            
            # 加總每個指標
            for key in totals:
                totals[key] += yearly_total[key]
                
            for key in direct_ghg_emission:
                direct_ghg_emission[key] += yearly_direct[key]
            
            for category in report['report_yearly_category']:
                if category['category_id'] in categories2_ids:
                    ghg_emission_tables['totalCategories2Value'] += category['totalGWPEmissionValue']
                    continue
                if category['category_id'] in categories3_ids:
                    ghg_emission_tables['totalCategories3Value'] += category['totalGWPEmissionValue']
                    continue
                if category['category_id'] in categories4_ids:
                    ghg_emission_tables['totalCategories4Value'] += category['totalGWPEmissionValue']
                    continue
                
                if category['category_id'] == stationary_id:
                    ghg_emission_tables['totalStationaryCombustionValue'] += category['totalGWPEmissionValue']
                elif category['category_id'] == mobile_id:
                    ghg_emission_tables['totalMobileCombustionValue'] += category['totalGWPEmissionValue']
                elif category['category_id'] == industrial_id:
                    ghg_emission_tables['totalIndustrialProcesses'] += category['totalGWPEmissionValue']
                elif category['category_id'] == refrigerants_id:
                    ghg_emission_tables['totalRefrigerantsFugitives'] += category['totalGWPEmissionValue']
                else:
                    pass
    
    return totals, ghg_emission_tables, direct_ghg_emission

# pipeline = [
    # # 展開 report_data 陣列
    # {"$unwind": "$report_data"},
    
    # # 匹配 year 為 2024 的文件
    # {"$match": {"report_data.year": 2024}},
    # ]

    # # 執行查詢
    # results1 = table.aggregate(pipeline)