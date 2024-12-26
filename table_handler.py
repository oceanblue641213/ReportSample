def organization_boundaries_setting_range_2_1_1(data):
    result = [(item.get('companyName'), item.get('address', "")) for item in data]
    
    return result

def ghg_emission_list_2_3_1(data):
    result = [(item.get('category'), item.get('fuelType', ""), item.get('co2EmissionValue', ""), item.get('ch4EmissionValue', ""), item.get('n2oEmissionValue', ""), item.get('hfcsEmissionValue', ""), item.get('pfcsEmissionValue', ""), item.get('sf6EmissionValue', ""), item.get('nf3EmissionValue', "")) for item in data]
    
    return result

def indirect_ghg_emission_evaluation_2_3_3(data, categories):
    """
    如果遵從要求(reqByLaw)==true，後面欄位都不用填，最後一欄是否量化計算填"V"
    如果遵德要求(reqByLaw)==false，後面欄位都要填，最後一欄是否量化計算填"X"
    """
    ids = [str(item['_id']) for item in categories]
    query = [item for item in data.get('rating') if str(item.get('subCategory')) in ids]
    threshold = data.get('threshold')
    result = [
                (
                    "是" if item.get('reqByLaw') == True else "否",
                    "" if item.get('reqByLaw') == True else ("" if item.get('rating1') is None or item.get('rating1') == "" else item.get('rating1')),
                    "" if item.get('reqByLaw') == True else ("" if item.get('rating2') is None or item.get('rating2') == "" else item.get('rating2')),
                    "" if item.get('reqByLaw') == True else ("" if item.get('rating3') is None or item.get('rating3') == "" else item.get('rating3')),
                    "" if item.get('reqByLaw') == True else ("" if item.get('rating4') is None or item.get('rating4') == "" else item.get('rating4')),
                    "" if item.get('reqByLaw') == True else ("" if item.get('rating5') is None or item.get('rating5') == "" else item.get('rating5')),
                    "" if item.get('reqByLaw') == True else int(sum(float(item.get(f'rating{i}')) if item.get(f'rating{i}') is not None and item.get(f'rating{i}') != "" else 0 for i in range(1, 6))),
                    "V" if (item.get('reqByLaw') == True or sum(float(item.get(f'rating{i}')) if item.get(f'rating{i}') is not None and item.get(f'rating{i}') != "" else 0 for i in range(1, 6)) >= threshold) else "X",
                ) for item in query
            ]
    
    return result

def direct_ghg_emission_3_3_1(data):
    # 首先過濾出 scope 為 1 的資料，並依照 scopeName 分組
    scope1_data = {}
    for asset in data:
        for item in asset['companyAsset']:
            if item['scope'] == '1':
                scopeName = item['scopeName']
                if scopeName not in scope1_data:
                    scope1_data[scopeName] = []
                scope1_data[scopeName].extend(item['assets'])

    # 再依照 activityName 分類，每個 activityName 只取一筆
    pre_result = {}
    for scopeName, assets in scope1_data.items():
        pre_result[scopeName] = {}
        seen_activities = set()
        
        for asset in assets:
            activity_name = asset.get('activityName', '')  # 有些資料可能沒有 activityName
            if activity_name and activity_name not in seen_activities:
                seen_activities.add(activity_name)
                pre_result[scopeName][activity_name] = asset.get('fuelType')
    
    result = []
    for scopeName, activities in pre_result.items():
        # 跳過空的活動設施
        if not activities:
            continue
            
        for activityName, fuelType in activities.items():
            # 如果 fuelType 是 None，轉換成空字串
            result.append((
                scopeName,
                activityName,
                "" if fuelType is None else fuelType
            ))
                
    return result



def indirect_ghg_emission_3_3_3(data):
    # 首先過濾出 scope 為 3或4 的資料，並依照 scopeName 分組
    scope_data = {}
    for asset in data:
        for item in asset['companyAsset']:
            if item['scope'] == '3' or item['scope'] == '4':
                scopeName = item['scopeName']
                if scopeName not in scope_data:
                    scope_data[scopeName] = []
                scope_data[scopeName].extend(item['assets'])

    # 再依照 activityName 分類，每個 activityName 只取一筆
    pre_result = {}
    for scopeName, assets in scope_data.items():
        pre_result[scopeName] = {}
        seen_activities = set()
        
        for asset in assets:
            activity_name = asset.get('activityName', '')  # 有些資料可能沒有 activityName
            if activity_name and activity_name not in seen_activities:
                seen_activities.add(activity_name)
                pre_result[scopeName][activity_name] = asset.get('fuelType')
    
    result = []
    for scopeName, activities in pre_result.items():
        # 跳過空的活動設施
        if not activities:
            continue
            
        for activityName, fuelType in activities.items():
            # 如果 fuelType 是 None，轉換成空字串
            result.append((
                scopeName,
                activityName,
                "" if fuelType is None else fuelType
            ))
                
    return result