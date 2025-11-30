# -*- coding: utf-8 -*-
from pyrevit import forms, revit
from Autodesk.Revit.DB import *
import os
import re
import time
import clr
import csv  # <--- Добавляем стандартную библиотеку для работы с CSV

# Пробуем подключить Excel
try:
    clr.AddReference("Microsoft.Office.Interop.Excel")
    from Microsoft.Office.Interop import Excel
except:
    Excel = None

# ==========================================================
# --- НАСТРОЙКИ ---
TARGET_SCHEDULE_NAME = "SHN_CommonBOQ" 
SERVER_ROOT_PATH = r"F:\REVIT_SHN\CHECK\Parameters_BOQ"
FILE_NAME_BASE = "SHN_CommonBOQ"
# ==========================================================

doc = revit.doc

def clean_filename(text):
    return re.sub(r'[\\/*?:"<>|]', '_', text).strip()

def get_export_folder():
    try:
        p_info = doc.ProjectInformation
        project_name = p_info.Name if p_info.Name else "Unassigned_Project"
        
        model_title = doc.Title
        if model_title.lower().endswith('.rvt'):
            model_title = model_title[:-4]
            
        safe_project = clean_filename(project_name)
        safe_model = clean_filename(model_title)
        
        full_path = os.path.join(SERVER_ROOT_PATH, safe_project, safe_model)
        
        if not os.path.exists(full_path):
            os.makedirs(full_path)
            
        return full_path
    except:
        return None

def csv_to_html(csv_path, html_path):
    """Конвертация CSV -> HTML (Исправленная логика)"""
    try:
        current_date = time.strftime("%Y-%m-%d %H:%M")
        
        # Читаем CSV правильно, используя библиотеку csv
        # Это решит проблему с запятыми внутри текста
        rows = []
        with open(csv_path, 'r') as f:
            # csv.reader автоматически обрабатывает кавычки и разделители
            reader = csv.reader(f)
            for row in reader:
                if row: # Пропускаем пустые строки
                    rows.append(row)

        if not rows:
            return False

        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body { font-family: Arial, sans-serif; padding: 20px; background-color: #fff; }
                h2 { text-align: center; margin-bottom: 5px; color: #333; }
                p.info { text-align: center; color: gray; font-size: 12px; margin-top: 0; margin-bottom: 20px; }
                table { 
                    width: 100%; 
                    border-collapse: collapse; 
                    font-size: 12px; 
                    box-shadow: 0 0 20px rgba(0, 0, 0, 0.15); 
                }
                th, td { 
                    border: 1px solid #dddddd; 
                    padding: 8px 12px; 
                    text-align: left; 
                    vertical-align: top;
                }
                th { 
                    background-color: #009879; 
                    color: #ffffff; 
                    font-weight: bold; 
                    position: sticky; top: 0; 
                }
                tr:nth-child(even) { background-color: #f3f3f3; }
                tr:hover { background-color: #f1f1f1; }
            </style>
        </head>
        <body>
            <h2>""" + TARGET_SCHEDULE_NAME + """</h2>
            <p class='info'>Model: """ + doc.Title + """ | Date: """ + current_date + """</p>
            <table>
        """
        
        for i, cells in enumerate(rows):
            html_content += "<tr>"
            tag = "th" if i == 0 else "td"
            
            for cell in cells:
                # Если ячейка пустая, ставим неразрывный пробел, чтобы рамка рисовалась
                cell_data = cell if cell.strip() != "" else "&nbsp;"
                
                # Декодируем текст, если он пришел в неправильной кодировке (актуально для старых движков)
                try:
                    cell_data = cell_data.decode('utf-8')
                except:
                    pass # Если это уже юникод или обычная строка, оставляем как есть
                
                html_content += "<{0}>{1}</{0}>".format(tag, cell_data)
            
            html_content += "</tr>"

        html_content += "</table></body></html>"
        
        # Записываем HTML
        import codecs
        with codecs.open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        return True
    except Exception as e:
        # Можно вывести ошибку в консоль для отладки: print(e)
        return False

def convert_csv_to_xlsx(csv_path, xlsx_path):
    """Конвертация CSV -> XLSX"""
    if not Excel: return False
    ex_app = None
    workbook = None
    try:
        ex_app = Excel.ApplicationClass()
        ex_app.Visible = False
        ex_app.DisplayAlerts = False
        
        # Format=6 (CSV), Delimiter="," - Excel сам разберется с кавычками лучше
        workbook = ex_app.Workbooks.Open(csv_path, Format=6, Delimiter=",")
        ws = workbook.Worksheets[1]

        ws.UsedRange.Columns.AutoFit()
        ws.Rows[1].Font.Bold = True
        ws.UsedRange.Borders.LineStyle = 1 

        if os.path.exists(xlsx_path):
            try: os.remove(xlsx_path)
            except: pass
            
        workbook.SaveAs(xlsx_path, 51) # 51 = xlOpenXMLWorkbook (xlsx)
        return True
    except: return False
    finally:
        if workbook: workbook.Close(SaveChanges=False)
        if ex_app: ex_app.Quit()
        try:
            import System.Runtime.InteropServices
            if workbook: System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            if ex_app: System.Runtime.InteropServices.Marshal.ReleaseComObject(ex_app)
        except: pass

def main():
    if not os.path.exists(SERVER_ROOT_PATH): return 
    export_folder = get_export_folder()
    if not export_folder: return

    target_view = None
    try:
        collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
        for v in collector:
            if v.Name == TARGET_SCHEDULE_NAME and not v.IsTemplate:
                target_view = v
                break
    except: return
    
    if not target_view: return

    # Пути к файлам
    filename_csv = FILE_NAME_BASE + ".csv"
    full_csv_path = os.path.join(export_folder, filename_csv)
    
    filename_html = FILE_NAME_BASE + ".html"
    full_html_path = os.path.join(export_folder, filename_html)
    
    filename_xlsx = FILE_NAME_BASE + ".xlsx"
    full_xlsx_path = os.path.join(export_folder, filename_xlsx)

    try:
        # 1. ЭКСПОРТ CSV
        opt = ViewScheduleExportOptions()
        opt.Title = False # Заголовок убираем, берем из имени спецификации
        opt.TextQualifier = ExportTextQualifier.DoubleQuote # Обязательно кавычки!
        opt.FieldDelimiter = ","
        
        # Revit 2022+ может требовать кодировку UTF-8, по умолчанию часто UTF-16
        # Но Python csv модуль лучше работает с чистым текстом.
        # Обычно экспорт Revit достаточно стабилен.
        
        if os.path.exists(full_csv_path):
            try: os.remove(full_csv_path)
            except: filename_csv = FILE_NAME_BASE + "_new.csv"
            full_csv_path = os.path.join(export_folder, filename_csv)
        
        target_view.Export(export_folder, filename_csv, opt)
        
        # Ждем, пока файл освободится
        max_retries = 10
        for _ in range(max_retries):
            if os.path.exists(full_csv_path):
                try:
                    with open(full_csv_path, 'r'): pass
                    break
                except: time.sleep(0.5)
            else:
                time.sleep(0.5)
        
        if not os.path.exists(full_csv_path): return

        # 2. ГЕНЕРАЦИЯ HTML (Использует CSV)
        csv_to_html(full_csv_path, full_html_path)

        # 3. ГЕНЕРАЦИЯ EXCEL (Использует CSV)
        convert_csv_to_xlsx(full_csv_path, full_xlsx_path)

        # 4. ВАЖНО: МЫ БОЛЬШЕ НЕ УДАЛЯЕМ CSV (как и просил)
        
    except:
        pass

main()