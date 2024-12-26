import upload_file as upf
from db_connection import create_connection as conn
import datetime
from dotenv import load_dotenv
from system_sub_categories import get_system_sub_categories_data
from site_modules import get_site_modules_data
from onbording_ratings import get_onbording_ratings_data
from company_assets import get_company_assets_data
from ghg_reports import get_ghg_reports_data
from table_handler import (
    organization_boundaries_setting_range_2_1_1, 
    ghg_emission_list_2_3_1, 
    indirect_ghg_emission_evaluation_2_3_3, 
    direct_ghg_emission_3_3_1,
    indirect_ghg_emission_3_3_3
    )
import os
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document

def main():
    # input_company_id = input('輸入公司id：')
    # input_year = input('輸入欲查詢年份：')
    company_id = "652cf86ce76f185628e48d9e" # 64a39d240db2a9a8f4a50893 # 63d0ac0a22d4cdf11191097e
    year = "2025"
    today = datetime.date.today()
    formatted_date = today.strftime("%Y年%m月%d日")
    
    load_dotenv()  # 讀取 .env 文件中的環境變數
    conn_string = os.getenv('DB_CONNECTIONSTRING')
    db_name = os.getenv('DB_NAME')
    
    with conn(conn_string, db_name) as db:
        try:
            all_collections = db.get_all_collections()
            site_modules, company_assets, onboarding_ratings, total_Emission, ghg_report, system_sub_categories, translations = (
            all_collections['site_modules'],
            all_collections['company_assets'],
            all_collections['onboarding_ratings'],
            all_collections['total_Emission'],
            all_collections['ghg_report'],
            all_collections['system_sub_categories'],
            all_collections['translations']
            )
            
            categories = get_system_sub_categories_data(system_sub_categories)
            all_site_modules_datas = get_site_modules_data(site_modules, company_id, fields=['companyName', 'levelStatus', 'usedRegion', 'address', 'EFGroup'])
            all_site_modules_ids = [item.get('_id') for item in all_site_modules_datas]
            sorted_data = sorted(all_site_modules_datas, key=lambda x: x.get('levelStatus') != 'parent')
            
            parent = next((item for item in all_site_modules_datas if item.get('levelStatus') == 'parent'), None)
            parent_rating = get_onbording_ratings_data(onboarding_ratings, parent.get('_id'))
            all_company_assets = get_company_assets_data(company_assets, all_site_modules_ids)
            all_ghg_reports_datas = get_ghg_reports_data(ghg_report, parent.get('_id'), fields=['emissionFactors']).get('emissionFactors')
            
            print("開始:")
            
            replacements = {
            '[company]': parent.get('companyName'),
            '[year]': parent_rating.get('baseYear'),
            '[report date]': formatted_date,
            '[covered_range_from_to]': 
                parent_rating.get('baseYearPeriodStart').strftime('%Y年%m月%d日')
                + '至'
                + parent_rating.get('baseYearPeriodEnd').strftime('%Y年%m月%d日'),
            '[onbording_ratings_type]': parent_rating.get('type'),
            '[threshold]': str(parent_rating.get('threshold')),
            }
            
            # 打開文檔
            doc = Document('溫室氣體盤查報告書v2.0.docx')
            
            # 替換文本
            upf.replace_text_in_word(doc, replacements)
            
            # 替換表格
            table00_data = organization_boundaries_setting_range_2_1_1(sorted_data)
            replace_table(doc.tables[0], table00_data)
            
            table01_data = ghg_emission_list_2_3_1(all_ghg_reports_datas)
            replace_table(doc.tables[1], table01_data, n = 2)
            
            table02_data = indirect_ghg_emission_evaluation_2_3_3(parent_rating, categories)
            replace_table(doc.tables[2], table02_data, n = 1, h = 2)
            
            table06_data = direct_ghg_emission_3_3_1(all_company_assets)
            replace_table(doc.tables[6], table06_data)
            
            table08_data = indirect_ghg_emission_3_3_3(all_company_assets)
            replace_table(doc.tables[8], table08_data)
            
            print("結束")
    
            # 保存新文檔
            doc.save('updated_document.docx')
            
        except Exception as e:
            print(f"發生錯誤：{e}")
    

def replace_table(table, new_data, n: int = 1, h: int = 0):
    """
    替換表格內容，保留標題行並根據數據動態調整表格大小
    
    Args:
        table: Word文檔中的表格對象
        new_data: 要插入的新數據（二維列表）
        n: 標題行的行數
        h: 要跳過的列數
    """
    # 獲取需要的行數（數據行數）
    required_rows = len(new_data) + n  # 因為要保留標題行
    current_rows = len(table.rows)
    
    # 調整表格大小
    if required_rows > current_rows:
        # 需要添加行
        for _ in range(required_rows - current_rows):
            table.add_row()
    elif required_rows < current_rows:
        # 需要刪除多餘的行
        for _ in range(current_rows - required_rows):
            table._element.remove(table.rows[-1]._element)
    
    # 然後填充數據行
    for i, row_data in enumerate(new_data, start=n):  # start=n 跳過標題行
        for j, cell_data in enumerate(row_data, start=h):
            if j < len(table.columns):  # 確保不超過列數
                cell = table.cell(i, j)
                cell.text = ""
                paragraph = cell.paragraphs[0]
                cell_data = str(cell_data)
                if cell_data == "True":
                    cell_data = "V"
                elif cell_data == "False":
                    cell_data = ""
                else:
                    pass
                run = paragraph.add_run(cell_data)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


if __name__ == '__main__':
    main()


    