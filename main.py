import upload_file as upf
from db_connection import create_connection as conn
import datetime
from dotenv import load_dotenv
from site_modules import get_site_modules_data
from onbording_ratings import get_onbording_ratings_data
from company_assets import get_company_assets_data
import os

def main():
    # input_company_id = input('輸入公司id：')
    # input_year = input('輸入欲查詢年份：')
    company_id = "64a39d240db2a9a8f4a50893" # 64a39d240db2a9a8f4a50893 # 63d0ac0a22d4cdf11191097e
    year = "2025"
    today = datetime.date.today()
    formatted_date = today.strftime("%Y年%m月%d日")
    
    load_dotenv()  # 讀取 .env 文件中的環境變數
    conn_string = os.getenv('DB_CONNECTIONSTRING')
    
    with conn(conn_string) as db:
        try:
            all_collections = db.get_all_collections()
            site_modules, company_assets, onboarding_ratings, total_Emission, translations = (
            all_collections['site_modules'],
            all_collections['company_assets'],
            all_collections['onboarding_ratings'],
            all_collections['total_Emission'],
            all_collections['translations']
            )
            
            all_site_modules = get_site_modules_data(site_modules, company_id)
            all_site_modules_ids = [item.get('_id') for item in all_site_modules]
            parent = next((item for item in all_site_modules if item.get('levelStatus') == 'parent'), None)
            
            parent_rating = get_onbording_ratings_data(onboarding_ratings, parent.get('_id'))
            all_company_assets = get_company_assets_data(company_assets, all_site_modules_ids)
            
            print("查詢結果:")
            
            replacements = {
            '[company]': parent.get('companyName'),
            '[year]': parent_rating.get('baseYear'),
            '[report date]': formatted_date,
            '[covered_range_from_to]': parent_rating.get('baseYearPeriodStart').strftime('%Y年%m月%d日') + '至' + parent_rating.get('baseYearPeriodEnd').strftime('%Y年%m月%d日'),
            '[onbording_ratings_type]': parent_rating.get('type'),
            '[threshold]': str(parent_rating.get('threshold')),
            }
            
            upf.replace_text_in_word('溫室氣體盤查報告書v2.0.docx', replacements)
            # upf.replace_text_in_word('updated_document.docx', replacements)
            
        except Exception as e:
            print(f"發生錯誤：{e}")
    
    



if __name__ == '__main__':
    main()


    