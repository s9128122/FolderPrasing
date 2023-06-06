import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
from openpyxl.utils import column_index_from_string

try:
    # 獲取當前腳本所在目錄
    current_directory = os.path.dirname(os.path.abspath(__file__))

    # 讀取文件夾和文件的階層結構
    folder_structure = {}

    def read_folder_structure(folder_path, level=0):
        folder_name = os.path.basename(folder_path)
        folder_structure[folder_path] = {'name': folder_name, 'level': level, 'files': [], 'subfolders': []}

        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isdir(item_path):
                folder_structure[folder_path]['subfolders'].append(item_path)
                read_folder_structure(item_path, level + 1)
            else:
                folder_structure[folder_path]['files'].append(item)

    # 讀取文件夾和文件的階層結構
    read_folder_structure(current_directory)

    # 創建txt文件，記錄文件夾和文件的階層結構
    text_file_path = os.path.join(current_directory, '所有資料夾檔案.txt')
    exclude_list=['所有資料夾檔案.txt', '所有資料夾檔案.xlsx', '列出資料名稱.py','列出資料名稱_fin.py','~$所有資料夾檔案.xlsx']
    def write_folder_structure(folder_info, indent=0, file=None):
        folder_name = folder_info['name']
        if folder_name not in exclude_list:
            file.write(' ' * indent + folder_name + '/' + '\n')
            files = folder_info['files']
            for file_name in files:
                if file_name not in exclude_list:
                    file.write(' ' * (indent + 4) + file_name + '\n')

            subfolders = folder_info['subfolders']
            for subfolder_path in subfolders:
                subfolder_info = folder_structure[subfolder_path]
                write_folder_structure(subfolder_info, indent + 4, file)

    # 輸出文件夾結構到txt
    with open(text_file_path, 'w', encoding='utf-8') as text_file:
        for folder_path, folder_info in folder_structure.items():
            if folder_info['level'] == 0 and folder_info['name'] != '列出資料名稱.py':
                write_folder_structure(folder_info, file=text_file)

    # 創建Excel
    excel_file_path = os.path.join(current_directory, '所有資料夾檔案.xlsx')
    workbook = Workbook()
    workbook.remove(workbook.active)  # 刪除default的工作表

    # 創建"All Folders"工作表，用於列出文件夾階層結構
    all_folders_sheet = workbook.create_sheet(title='All Folders')

    # 寫入文件夾階層結構和文件名
    row = 1
    for folder_path, folder_info in folder_structure.items():
        folder_name = folder_info['name']
        level = folder_info['level']
        if folder_name not in exclude_list:
            folder_name_cell = all_folders_sheet.cell(row=row, column=level + 1)
            folder_name_cell.value = folder_name
            folder_sheet_name = re.sub(r'[\/:*?\\[\]]', '_', folder_name)[:30]
            folder_name_cell.hyperlink = f"所有資料夾檔案.xlsx#'{folder_sheet_name}'!A1"

            # 獲取該文件夾內的文件名
            files = folder_info['files']
            if files:
                for file_name in files:
                    if file_name not in exclude_list:
                        file_name_cell = all_folders_sheet.cell(row=row + 1, column=6)
                        file_path = os.path.join(folder_path, file_name)
                        file_hyperlink = f"file:///{file_path}"
                        file_name_cell.hyperlink = file_hyperlink
                        file_name_cell.value = file_name

                        row += 1

            row += 1

            folder_sheet = workbook.create_sheet(title=folder_sheet_name)
            folder_sheet['A1'] = folder_sheet_name
            folder_sheet.column_dimensions['A'].width = 50
            folder_sheet['A1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            subfolders = folder_info['subfolders']
            files = folder_info['files']

            if not subfolders and not files:
                folder_sheet['A2'] = '找不到子文件夾或文件。'
            else:
                subfolder_row = 2
                for subfolder_path in subfolders:
                    subfolder_name = os.path.basename(subfolder_path)
                    subfolder_sheet_name = re.sub(r'[\/:*?\\[\]]', '_', subfolder_name)[:30]
                    folder_sheet.cell(row=subfolder_row, column=1).value = subfolder_name
                    folder_sheet.cell(row=subfolder_row, column=1).hyperlink = f"所有資料夾檔案.xlsx#'{subfolder_sheet_name}'!A1"
                    subfolder_row += 1

                file_row = subfolder_row
                for file_name in files:
                    if file_name not in exclude_list:
                        file_path = os.path.join(folder_path, file_name)
                        file_hyperlink = f"file:///{file_path}"
                        folder_sheet.cell(row=file_row, column=2).hyperlink = file_hyperlink
                        folder_sheet.cell(row=file_row, column=2).value = file_name
                        file_row += 1

    # 調整總表的列寬
    all_folders_sheet.column_dimensions[get_column_letter(all_folders_sheet.max_column)].width = 30

    # 在每個工作表中添加超鏈接到第一張工作表
    for sheet_name in workbook.sheetnames:
        if sheet_name != 'All Folders':
            sheet = workbook[sheet_name]
            sheet['A1'].hyperlink = "所有資料夾檔案.xlsx#'All Folders'!A1"
            sheet['A1'].value = sheet['A1'].value + ' (返回總表)'

    # 在每個工作表中添加超鏈接樣式
    hyperlink_style = Font(underline='single', color='0563C1')
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.hyperlink:
                    cell.font = hyperlink_style
    
    # 根據F欄的連續值 分組並收合
    def group_continuous_rows(sheet):
        max_row = sheet.max_row
        group_start = None
        previous_value = None
        for row in range(2, max_row + 1):
            current_value = sheet.cell(row=row, column=6).value
    
            if current_value is not None:
                if group_start is None:
                    group_start = row
            else:
                if group_start is not None:
                    sheet.row_dimensions.group(group_start, row - 1, hidden=True)
                    group_start = None
    
    
    # 對"All Folders"工作表進行分組並收合
    group_continuous_rows(all_folders_sheet)
    
    # 對每個文件夾的工作表進行分組並收合
    for sheet_name in workbook.sheetnames:
        if sheet_name != 'All Folders':
            sheet = workbook[sheet_name]
            group_continuous_rows(sheet)



    # 保存Excel文件
    workbook.save(excel_file_path)

except Exception as e:
    print(f'錯誤: {str(e)}')

print(f'文件夾和文件的階層結構已匯出至文本文件: {text_file_path}')
print(f'文件夾和文件的階層結構已匯出至Excel文件: {excel_file_path}')
