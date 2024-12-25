import upload_file as upf
from db_connection import create_connection as conn
import datetime
from dotenv import load_dotenv
from site_modules import get_site_modules_data
from onbording_ratings import get_onbording_ratings_data
from company_assets import get_company_assets_data
import os
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document

def main():
    # input_company_id = input('輸入公司id：')
    # input_year = input('輸入欲查詢年份：')
    company_id = "64a39d240db2a9a8f4a50893" # 64a39d240db2a9a8f4a50893 # 63d0ac0a22d4cdf11191097e
    year = "2025"
    today = datetime.date.today()
    formatted_date = today.strftime("%Y年%m月%d日")
    
    load_dotenv()  # 讀取 .env 文件中的環境變數
    conn_string = os.getenv('DB_CONNECTIONSTRING')
    db_name = os.getenv('DB_NAME')
    
    with conn(conn_string, db_name) as db:
        try:
            all_collections = db.get_all_collections()
            site_modules, company_assets, onboarding_ratings, total_Emission, translations = (
            all_collections['site_modules'],
            all_collections['company_assets'],
            all_collections['onboarding_ratings'],
            all_collections['total_Emission'],
            all_collections['translations']
            )
            
            all_site_modules_datas = get_site_modules_data(site_modules, company_id, fields=['companyName', 'levelStatus', 'usedRegion', 'address', 'EFGroup'])
            all_site_modules_ids = [item.get('_id') for item in all_site_modules_datas]
            sorted_data = sorted(all_site_modules_datas, 
                                key=lambda x: x.get('levelStatus') != 'parent')
            
            parent = next((item for item in all_site_modules_datas if item.get('levelStatus') == 'parent'), None)
            parent_rating = get_onbording_ratings_data(onboarding_ratings, parent.get('_id'))
            all_company_assets = get_company_assets_data(company_assets, all_site_modules_ids)
            
            print("開始:")
            
            replacements = {
            '[company]': parent.get('companyName'),
            '[year]': parent_rating.get('baseYear'),
            '[report date]': formatted_date,
            '[covered_range_from_to]': parent_rating.get('baseYearPeriodStart').strftime('%Y年%m月%d日') + '至' + parent_rating.get('baseYearPeriodEnd').strftime('%Y年%m月%d日'),
            '[onbording_ratings_type]': parent_rating.get('type'),
            '[threshold]': str(parent_rating.get('threshold')),
            }
            
            # 打開文檔
            doc = Document('溫室氣體盤查報告書v2.0.docx')
    
            upf.replace_text_in_word(doc, replacements)
            # upf.replace_text_in_word('updated_document.docx', replacements)
            
            table01_data = [(item.get('companyName'), item.get('address', "")) for item in sorted_data]
            
            replace_table(doc.tables[0], table01_data)
            
            print("結束")
    
            # 保存新文檔
            doc.save('updated_document.docx')
            
        except Exception as e:
            print(f"發生錯誤：{e}")
    

def replace_table(table, new_data):
    """
    替換表格內容，保留標題行並根據數據動態調整表格大小
    
    Args:
        table: Word文檔中的表格對象
        new_data: 要插入的新數據（二維列表）
    """
    # 保存標題行
    # header_cells = []
    # for cell in table.rows[0].cells:
    #     header_cells.append(cell.text)
    
    # 獲取需要的行數（數據行數）
    required_rows = len(new_data) + 1  # +1 是因為要保留標題行
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
    
    # 重新填充數據
    # 首先恢復標題行
    # for j, header_text in enumerate(header_cells):
    #     cell = table.cell(0, j)
    #     cell.text = ""
    #     paragraph = cell.paragraphs[0]
    #     run = paragraph.add_run(str(header_text))
    #     paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 然後填充數據行
    for i, row_data in enumerate(new_data, start=1):  # start=1 跳過標題行
        for j, cell_data in enumerate(row_data):
            if j < len(table.columns):  # 確保不超過列數
                cell = table.cell(i, j)
                cell.text = ""
                paragraph = cell.paragraphs[0]
                run = paragraph.add_run(str(cell_data))
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER



if __name__ == '__main__':
    main()


    