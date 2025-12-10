# -*- coding: utf-8 -*-
from pyrevit import revit
from Autodesk.Revit.DB import *
import os
import re
import time
import clr
import sys

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

# Имя столбца, по которому фильтруем Family
FILTER_FAMILY_COLUMN_NAME   = u"Family"    # должен совпадать с заголовком в спецификации
# Имя столбца, по которому фильтруем Category
FILTER_CATEGORY_COLUMN_NAME = u"Category"  # ← ПОМЕНЯЙ на фактическое имя колонки, например u"Revit Category"
# ==========================================================

doc = revit.doc


def clean_filename(text):
    return re.sub(r'[\\/*?:"<>|]', '_', text).strip()


def get_export_folder():
    """Формируем путь: ROOT\\ProjectName\\ModelName и создаём, если нет."""
    try:
        p_info = doc.ProjectInformation
        project_name = p_info.Name if p_info and p_info.Name else "Unassigned_Project"

        model_title = doc.Title or "Unnamed_Model"
        if model_title.lower().endswith('.rvt'):
            model_title = model_title[:-4]

        safe_project = clean_filename(project_name)
        safe_model = clean_filename(model_title)

        full_path = os.path.join(SERVER_ROOT_PATH, safe_project, safe_model)

        if not os.path.exists(full_path):
            os.makedirs(full_path)

        return full_path
    except Exception:
        return None


# ---------- CSV / HTML ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----------

def read_csv_text(csv_path):
    """Чтение CSV-файла с попыткой угадать кодировку (Revit часто даёт UTF-16)."""
    try:
        with open(csv_path, 'rb') as fb:
            raw = fb.read()
    except:
        return None

    for enc in ('utf-8-sig', 'utf-16', 'cp1255', 'cp1251'):
        try:
            return raw.decode(enc)
        except:
            continue

    try:
        return raw.decode(errors='ignore')
    except:
        return None


def parse_csv_line(line, delimiter=",", quote='"'):
    """
    Простой CSV-парсер:
    - понимает кавычки
    - запятые внутри кавычек не считаются разделителем
    - "" внутри строки -> одна "
    """
    cells = []
    current = []
    in_quotes = False
    i = 0
    length = len(line)

    while i < length:
        ch = line[i]

        if ch == quote:
            if in_quotes and i + 1 < length and line[i + 1] == quote:
                current.append(quote)
                i += 1
            else:
                in_quotes = not in_quotes
        elif ch == delimiter and not in_quotes:
            cells.append(u"".join(current))
            current = []
        else:
            current.append(ch)
        i += 1

    cells.append(u"".join(current))
    return cells


def html_escape(s):
    """Простое экранирование спецсимволов для HTML."""
    if s is None:
        return u""
    if isinstance(s, bytes):
        try:
            s = s.decode('utf-8')
        except:
            try:
                s = s.decode('cp1251')
            except:
                s = s.decode(errors='ignore')

    s = s.replace(u"&", u"&amp;")
    s = s.replace(u"<", u"&lt;")
    s = s.replace(u">", u"&gt;")
    s = s.replace(u"\"", u"&quot;")
    s = s.replace(u"'", u"&#39;")
    return s


def csv_to_html(csv_path, html_path):
    """Конвертация CSV -> HTML + фильтры по Family и Category."""
    try:
        text = read_csv_text(csv_path)
        if not text:
            raise Exception("Could not read CSV text")

        lines = text.splitlines()
        rows = []
        for line in lines:
            if line.strip() == "":
                continue
            rows.append(parse_csv_line(line))

        if not rows:
            rows = [[u"NO DATA"]]

        header = rows[0]

        # --- Поиск индекса колонки Family ---
        fam_index = -1
        target_family_header = (FILTER_FAMILY_COLUMN_NAME or u"").strip().lower()
        for i, h in enumerate(header):
            hn = (h or u"").strip().lower()
            if hn == target_family_header:
                fam_index = i
                break

        # --- Поиск индекса колонки Category (по имени из настройки) ---
        cat_index = -1
        target_cat_header = (FILTER_CATEGORY_COLUMN_NAME or u"").strip().lower()
        if target_cat_header:
            for i, h in enumerate(header):
                hn = (h or u"").strip().lower()
                if hn == target_cat_header:
                    cat_index = i
                    break

        # Список уникальных значений для фильтров
        family_values = []
        if fam_index >= 0:
            seen = set()
            for r in rows[1:]:
                if fam_index < len(r):
                    val = (r[fam_index] or u"").strip()
                    if val and val not in seen:
                        seen.add(val)
                        family_values.append(val)
            family_values.sort()

        category_values = []
        if cat_index >= 0:
            seen = set()
            for r in rows[1:]:
                if cat_index < len(r):
                    val = (r[cat_index] or u"").strip()
                    if val and val not in seen:
                        seen.add(val)
                        category_values.append(val)
            category_values.sort()

        current_date = time.strftime("%Y-%m-%d %H:%M")

        # --- HTML HEADER + CSS ---
        html_content = u"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background-color: #fff; }
        h2 { text-align: center; margin-bottom: 5px; color: #333; }
        p.info { text-align: center; color: gray; font-size: 12px; margin-top: 0; margin-bottom: 20px; }
        table.data-table {
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

        .filter-bar { margin: 10px 0 15px 0; }
        .dropdown { position: relative; display: inline-block; margin-right: 10px; }
        .dropbtn {
            padding: 6px 10px;
            border: 1px solid #ccc;
            background-color: #f8f8f8;
            cursor: pointer;
            font-size: 12px;
        }
        .dropdown-content {
            display: none;
            position: absolute;
            background-color: #ffffff;
            min-width: 220px;
            border: 1px solid #ccc;
            padding: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            z-index: 100;
            max-height: 300px;
            overflow-y: auto;
        }
        .dropdown:hover .dropdown-content { display: block; }
        .dropdown-content label {
            display: block;
            font-size: 12px;
            margin-bottom: 2px;
            cursor: pointer;
        }
        .filter-actions {
            font-size: 11px;
            margin-bottom: 4px;
        }
        .filter-actions a {
            cursor: pointer;
            text-decoration: underline;
        }
    </style>
</head>
<body>
"""

        # --- ЗАГОЛОВОК И ИНФО ---
        html_content += (
            u'    <h2>' + html_escape(TARGET_SCHEDULE_NAME) + u'</h2>\n'
            u'    <p class="info">Model: ' + html_escape(doc.Title or u"") +
            u' | Date: ' + html_escape(current_date) + u'</p>\n'
        )

        # --- ПАНЕЛЬ ФИЛЬТРОВ ---
        if (fam_index >= 0 and family_values) or (cat_index >= 0 and category_values):
            html_content += u'    <div class="filter-bar">\n'

            # Filter by Family
            if fam_index >= 0 and family_values:
                html_content += u"""        <div class="dropdown">
            <button class="dropbtn">Filter by """ + html_escape(FILTER_FAMILY_COLUMN_NAME) + u"""</button>
            <div class="dropdown-content family-filter">
                <div class="filter-actions">
                    <a id="fam-select-all">Select all</a> |
                    <a id="fam-select-none">Clear all</a>
                </div>
"""
                for val in family_values:
                    esc_val = html_escape(val)
                    html_content += (
                        u'                <label><input type="checkbox" value="' +
                        esc_val + u'" checked> ' + esc_val + u'</label>\n'
                    )

                html_content += u"""            </div>
        </div>
"""

            # Filter by Category
            if cat_index >= 0 and category_values:
                html_content += u"""        <div class="dropdown">
            <button class="dropbtn">Filter by """ + html_escape(FILTER_CATEGORY_COLUMN_NAME) + u"""</button>
            <div class="dropdown-content category-filter">
                <div class="filter-actions">
                    <a id="cat-select-all">Select all</a> |
                    <a id="cat-select-none">Clear all</a>
                </div>
"""
                for val in category_values:
                    esc_val = html_escape(val)
                    html_content += (
                        u'                <label><input type="checkbox" value="' +
                        esc_val + u'" checked> ' + esc_val + u'</label>\n'
                    )

                html_content += u"""            </div>
        </div>
"""

            html_content += u'    </div>\n'

        # --- ТАБЛИЦА ---
        html_content += u'    <table class="data-table">\n'
        # header
        html_content += u'    <thead>\n        <tr>\n'
        for cell in header:
            cell_data = cell if cell is not None else u""
            cell_data = html_escape(cell_data.strip())
            if cell_data == u"":
                cell_data = u"&nbsp;"
            html_content += u'            <th>' + cell_data + u'</th>\n'
        html_content += u'        </tr>\n    </thead>\n'

        # body
        html_content += u'    <tbody>\n'
        for row in rows[1:]:
            fam_val = u""
            if fam_index >= 0 and fam_index < len(row):
                val = row[fam_index] or u""
                fam_val = val.strip()

            cat_val = u""
            if cat_index >= 0 and cat_index < len(row):
                val = row[cat_index] or u""
                cat_val = val.strip()

            tr_open = u'        <tr class="data-row" data-family="' + html_escape(fam_val) + \
                      u'" data-category="' + html_escape(cat_val) + u'">'
            html_content += tr_open

            for j, cell in enumerate(row):
                cell_data = cell if cell is not None else u""
                if isinstance(cell_data, bytes):
                    try:
                        cell_data = cell_data.decode('utf-8')
                    except:
                        try:
                            cell_data = cell_data.decode('cp1251')
                        except:
                            cell_data = cell_data.decode(errors='ignore')

                cell_data = cell_data.strip()
                if cell_data == u"":
                    cell_data = u"&nbsp;"
                else:
                    cell_data = html_escape(cell_data)

                html_content += u"<td>" + cell_data + u"</td>"

            html_content += u"</tr>\n"

        html_content += u'    </tbody>\n</table>\n'

        # --- JS ДЛЯ ОБОИХ ФИЛЬТРОВ ---
        html_content += u"""
    <script>
    document.addEventListener('DOMContentLoaded', function() {
        var famCheckboxes = document.querySelectorAll('.family-filter input[type="checkbox"]');
        var catCheckboxes = document.querySelectorAll('.category-filter input[type="checkbox"]');

        function getActiveValues(nodeList) {
            var result = [];
            if (!nodeList) return result;
            for (var i = 0; i < nodeList.length; i++) {
                if (nodeList[i].checked) {
                    result.push(nodeList[i].value);
                }
            }
            return result;
        }

        function updateVisibility() {
            var activeFam = getActiveValues(famCheckboxes);
            var activeCat = getActiveValues(catCheckboxes);

            var rows = document.querySelectorAll('table.data-table tbody tr.data-row');
            for (var j = 0; j < rows.length; j++) {
                var row = rows[j];
                var vFam = row.getAttribute('data-family') || '';
                var vCat = row.getAttribute('data-category') || '';

                var famOK = (activeFam.length === 0) || (activeFam.indexOf(vFam) !== -1);
                var catOK = (activeCat.length === 0) || (activeCat.indexOf(vCat) !== -1);

                if (famOK && catOK) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            }
        }

        // события family
        for (var i = 0; i < famCheckboxes.length; i++) {
            famCheckboxes[i].addEventListener('change', updateVisibility);
        }
        var famAll = document.getElementById('fam-select-all');
        var famNone = document.getElementById('fam-select-none');
        if (famAll) {
            famAll.addEventListener('click', function(e) {
                e.preventDefault();
                for (var i = 0; i < famCheckboxes.length; i++) {
                    famCheckboxes[i].checked = true;
                }
                updateVisibility();
            });
        }
        if (famNone) {
            famNone.addEventListener('click', function(e) {
                e.preventDefault();
                for (var i = 0; i < famCheckboxes.length; i++) {
                    famCheckboxes[i].checked = false;
                }
                updateVisibility();
            });
        }

        // события category
        for (var i = 0; i < catCheckboxes.length; i++) {
            catCheckboxes[i].addEventListener('change', updateVisibility);
        }
        var catAll = document.getElementById('cat-select-all');
        var catNone = document.getElementById('cat-select-none');
        if (catAll) {
            catAll.addEventListener('click', function(e) {
                e.preventDefault();
                for (var i = 0; i < catCheckboxes.length; i++) {
                    catCheckboxes[i].checked = true;
                }
                updateVisibility();
            });
        }
        if (catNone) {
            catNone.addEventListener('click', function(e) {
                e.preventDefault();
                for (var i = 0; i < catCheckboxes.length; i++) {
                    catCheckboxes[i].checked = false;
                }
                updateVisibility();
            });
        }

        updateVisibility();
    });
    </script>
"""

        html_content += u"</body></html>"

        import codecs
        with codecs.open(html_path, 'w', encoding='utf-8') as hf:
            hf.write(html_content)

        return True

    except Exception as e:
        try:
            import codecs
            err_html = u"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body>
<p style="color:red;">Error while generating HTML from CSV.</p>
<p>{err}</p>
</body></html>""".format(err=html_escape(unicode(e)))
            with codecs.open(html_path, 'w', encoding='utf-8') as hf:
                hf.write(err_html)
        except:
            pass
        return False


def convert_csv_to_xlsx(csv_path, xlsx_path):
    """Конвертация CSV -> XLSX через Excel Interop."""
    if not Excel:
        return False

    ex_app = None
    workbook = None
    try:
        ex_app = Excel.ApplicationClass()
        ex_app.Visible = False
        ex_app.DisplayAlerts = False

        workbook = ex_app.Workbooks.Open(csv_path, Format=6, Delimiter=",")
        ws = workbook.Worksheets[1]

        ws.UsedRange.Columns.AutoFit()
        ws.Rows[1].Font.Bold = True
        ws.UsedRange.Borders.LineStyle = 1

        if os.path.exists(xlsx_path):
            try:
                os.remove(xlsx_path)
            except:
                pass

        workbook.SaveAs(xlsx_path, 51)  # 51 = xlOpenXMLWorkbook (xlsx)
        return True
    except Exception:
        return False
    finally:
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
        except:
            pass
        try:
            if ex_app:
                ex_app.Quit()
        except:
            pass

        try:
            import System.Runtime.InteropServices
            if workbook:
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            if ex_app:
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ex_app)
        except:
            pass


def main():
    if not os.path.exists(SERVER_ROOT_PATH):
        return

    export_folder = get_export_folder()
    if not export_folder:
        return

    target_view = None
    try:
        collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
        for v in collector:
            if not v.IsTemplate and v.Name == TARGET_SCHEDULE_NAME:
                target_view = v
                break
    except Exception:
        return

    if not target_view:
        return

    filename_csv = FILE_NAME_BASE + ".csv"
    full_csv_path = os.path.join(export_folder, filename_csv)

    filename_html = FILE_NAME_BASE + ".html"
    full_html_path = os.path.join(export_folder, filename_html)

    filename_xlsx = FILE_NAME_BASE + ".xlsx"
    full_xlsx_path = os.path.join(export_folder, filename_xlsx)

    try:
        opt = ViewScheduleExportOptions()
        opt.Title = False
        opt.TextQualifier = ExportTextQualifier.DoubleQuote
        opt.FieldDelimiter = ","

        if os.path.exists(full_csv_path):
            try:
                os.remove(full_csv_path)
            except:
                filename_csv = FILE_NAME_BASE + "_new.csv"
                full_csv_path = os.path.join(export_folder, filename_csv)

        target_view.Export(export_folder, filename_csv, opt)

        max_retries = 10
        for _ in range(max_retries):
            if os.path.exists(full_csv_path):
                try:
                    with open(full_csv_path, 'rb'):
                        pass
                    break
                except:
                    time.sleep(0.5)
            else:
                time.sleep(0.5)

        if not os.path.exists(full_csv_path):
            return

        csv_to_html(full_csv_path, full_html_path)
        convert_csv_to_xlsx(full_csv_path, full_xlsx_path)

    except Exception:
        pass


main()
