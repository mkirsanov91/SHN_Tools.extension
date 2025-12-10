# -*- coding: utf-8 -*-
"""
SHN Family Browser
------------------
Быстрый просмотр утверждённой библиотеки семейств с категориями, превью и Description.

 - Семейства берутся из F:\REVIT_SHN\SHN_Familys\test
 - Для каждого семейства ищется PNG в папке family_previews по маске:
        <ИмяСемейства>*.png
   (подойдут файлы типа 'SHN_Family - Floor Plan 0000.png')
 - Если превью не найдено, скрипт:
        * открывает семейство
        * ищет 3D-вид (если есть) или любой печатаемый вид
        * экспортирует PNG через ImageExportOptions в папку family_previews
 - В индекс добавляется параметр Description (BuiltIn ALL_MODEL_DESCRIPTION),
   если он есть у семейства/типа.
 - В окне:
        * слева — иконка
        * посередине — имя семейства
        * справа — Description
        * сверху — выбор категории + строка поиска по имени/Description
        * категория "(All categories)" включает все семейства.
"""

__title__ = 'Family\nBrowser'
__doc__ = 'Browse approved SHN families with preview thumbnails and Description, and load them into the project.'

from pyrevit import revit, forms
import clr
import os
import json
import System

# ======================================================================
# Revit API
# ======================================================================
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import Autodesk.Revit.DB as DB

# WPF
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Core')
from System.Runtime.CompilerServices import StrongBox
from System.Windows import Thickness
from System.Windows.Controls import StackPanel, Image, TextBlock, ListBoxItem
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri, UriKind

from System.Collections.Generic import List as Clist

doc = revit.doc
app = doc.Application

# ======================================================================
# ПУТИ
# ======================================================================

# Корень библиотеки утверждённых семейств
FAMILIES_ROOT = r"F:\REVIT_SHN\SHN_Familys\test"

# Папка кнопки (bundle)
BUTTON_DIR = os.path.dirname(__file__)

# Папка для превьюшек
PREVIEW_ROOT = os.path.join(BUTTON_DIR, 'family_previews')
if not os.path.exists(PREVIEW_ROOT):
    os.makedirs(PREVIEW_ROOT)

# Путь к JSON-индексу
INDEX_PATH = os.path.join(BUTTON_DIR, 'family_index.json')

# Размер картинки при экспорте
PREVIEW_PIXEL_SIZE = 256


# ======================================================================
# СЛУЖЕБНЫЙ КЛАСС ДЛЯ ЗАГРУЗКИ СЕМЕЙСТВ
# ======================================================================

class SimpleFamilyLoadOptions(DB.IFamilyLoadOptions):
    """Опции загрузки семей: всегда загружать, overwrite по выбору пользователя."""
    def __init__(self, overwrite=False):
        self._overwrite = overwrite

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = self._overwrite
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = self._overwrite
        return True


# ======================================================================
# ПОИСК / ЭКСПОРТ ПРЕВЬЮ
# ======================================================================

def _find_existing_preview(family_name):
    """Ищет в PREVIEW_ROOT любой PNG/JPG, имя которого начинается с имени семейства."""
    if not os.path.exists(PREVIEW_ROOT):
        return None

    for root, dirs, files in os.walk(PREVIEW_ROOT):
        for fname in files:
            lower = fname.lower()
            if not (lower.endswith('.png') or lower.endswith('.jpg') or lower.endswith('.jpeg')):
                continue
            if not fname.startswith(family_name):
                continue
            return os.path.join(root, fname)
    return None


def _get_preview_view_id(fam_doc):
    """Возвращает ElementId вида для превью: сначала 3D, потом любой печатаемый вид."""
    # 1) пробуем найти нормальный 3D-вид
    views3d = DB.FilteredElementCollector(fam_doc) \
                .OfClass(DB.View3D) \
                .WhereElementIsNotElementType() \
                .ToElements()

    for v3d in views3d:
        if v3d.IsTemplate:
            continue
        if hasattr(v3d, 'IsPerspective') and v3d.IsPerspective:
            continue
        if not v3d.CanBePrinted:
            continue
        return v3d.Id

    # 2) если нормального 3D нет — любой печатаемый вид (план/разрез)
    views = DB.FilteredElementCollector(fam_doc) \
              .OfClass(DB.View) \
              .WhereElementIsNotElementType() \
              .ToElements()

    for v in views:
        if v.IsTemplate:
            continue
        if not v.CanBePrinted:
            continue
        return v.Id

    # совсем ничего не нашли
    return None


def _get_family_description(fam_doc, owner_fam):
    """Пытается достать Description (ALL_MODEL_DESCRIPTION) у OwnerFamily или первого типа."""
    description_val = None
    try:
        # 1) Description у OwnerFamily
        if owner_fam:
            p = owner_fam.get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
            if p and p.HasValue:
                description_val = p.AsString()

        # 2) если не найдено — пробуем первый тип семейства
        if not description_val:
            sym = DB.FilteredElementCollector(fam_doc).OfClass(DB.FamilySymbol).FirstElement()
            if sym:
                p2 = sym.get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
                if p2 and p2.HasValue:
                    description_val = p2.AsString()
    except:
        pass

    return description_val


def _build_family_info(family_path):
    """
    Открывает семейный документ, читает категорию и Description,
    при необходимости экспортирует PNG из подходящего вида.
    Возвращает dict: name, category, path, preview (полный путь к файлу), description.
    """
    fam_doc = None
    category_name = "Unknown"
    description_val = None

    family_name = os.path.splitext(os.path.basename(family_path))[0]

    # сначала пробуем найти уже существующую картинку по имени семейства
    preview_path = _find_existing_preview(family_name)

    try:
        fam_doc = app.OpenDocumentFile(family_path)

        # --- категория семейства + Description ---
        owner_fam = fam_doc.OwnerFamily
        if owner_fam and owner_fam.FamilyCategory:
            category_name = owner_fam.FamilyCategory.Name

        description_val = _get_family_description(fam_doc, owner_fam)

        # --- если превьюшки нет, пробуем сделать экспорт ---
        if preview_path is None:
            try:
                view_id = _get_preview_view_id(fam_doc)

                if view_id:
                    base_no_ext = os.path.join(PREVIEW_ROOT, family_name)

                    opts = DB.ImageExportOptions()
                    opts.ExportRange = DB.ExportRange.SetOfViews
                    opts.HLRandWFViewsFileType = DB.ImageFileType.PNG
                    opts.FilePath = base_no_ext
                    opts.PixelSize = PREVIEW_PIXEL_SIZE
                    opts.ImageResolution = DB.ImageResolution.DPI_72

                    id_list = Clist[DB.ElementId]()
                    id_list.Add(view_id)
                    opts.SetViewsAndSheets(id_list)

                    fam_doc.ExportImage(opts)

                    # после экспорта ещё раз ищем любой файл с префиксом family_name
                    preview_path = _find_existing_preview(family_name)
                    if preview_path is None:
                        print("Image export did not create file for: {}".format(family_path))
                else:
                    print("No printable view found in family: {}".format(family_path))

            except Exception as ex2:
                print("Preview export error for {}: {}".format(family_path, ex2))

    except Exception as ex:
        print("Error reading family: {}\n{}".format(family_path, ex))
    finally:
        if fam_doc:
            fam_doc.Close(False)

    return {
        'name': family_name,
        'category': category_name,
        'path': family_path,
        'preview': preview_path if preview_path and os.path.exists(preview_path) else None,
        'description': description_val or ""
    }


# ======================================================================
# СБОР БИБЛИОТЕКИ / КЭШ
# ======================================================================

def _scan_library():
    """Обходит корень библиотеки и собирает информацию обо всех RFA."""
    if not os.path.exists(FAMILIES_ROOT):
        forms.alert(u"Family library folder not found:\n{}".format(FAMILIES_ROOT), exitscript=True)

    families = []

    for dirpath, dirnames, filenames in os.walk(FAMILIES_ROOT):
        for fname in filenames:
            if not fname.lower().endswith('.rfa'):
                continue

            full_path = os.path.join(dirpath, fname)
            info = _build_family_info(full_path)
            families.append(info)

    if not families:
        forms.alert(u"No RFA files found in:\n{}".format(FAMILIES_ROOT), exitscript=True)

    return families


def build_index(show_alert=True):
    """Полное перестроение индекса (медленно, но редко)."""
    if show_alert:
        forms.alert(
            u"Building SHN family library index.\n"
            u"This may take some time on first run.",
            ok=True
        )

    families = _scan_library()

    with open(INDEX_PATH, 'w') as f:
        json.dump(families, f)

    return families


def load_index():
    """Загрузка индекса из JSON. Если нет файла — возвращает None."""
    if not os.path.exists(INDEX_PATH):
        return None

    try:
        with open(INDEX_PATH, 'r') as f:
            families = json.load(f)
        return families
    except Exception as ex:
        print("Error loading index: {}".format(ex))
        return None


# ======================================================================
# WPF ОКНО
# ======================================================================

from pyrevit import forms as pyforms
import System.Windows


from pyrevit import forms as pyforms
import System.Windows


class FamilyBrowserWindow(pyforms.WPFWindow):
    def __init__(self, families):
        xaml_path = os.path.join(BUTTON_DIR, 'FamilyBrowser.xaml')
        pyforms.WPFWindow.__init__(self, xaml_path)

        self.families = families or []
        self._rebuild_categories()

    # ------------------------------ helpers ----------------------------

    def _rebuild_categories(self):
        cats = sorted(set(f.get('category', '') for f in self.families if f.get('category')))
        # добавляем вариант "All categories"
        self.categories = ['(All categories)'] + cats
        self.categoryCombo.ItemsSource = self.categories
        if self.categories:
            self.categoryCombo.SelectedIndex = 0
        self._populate_family_list()

    def _populate_family_list(self):
        """Заполняет список семейств с учётом категории и поиска."""
        self.familyList.Items.Clear()

        # текущая категория
        current_cat = self.categoryCombo.SelectedItem
        if current_cat is not None:
            current_cat = str(current_cat)
        else:
            current_cat = '(All categories)'

        # текст поиска
        query = ""
        if hasattr(self, 'searchBox') and self.searchBox.Text:
            query = self.searchBox.Text.strip().lower()

        for fam in self.families:
            cat = fam.get('category', '')
            name = fam.get('name', '')
            desc = fam.get('description', '') or ''

            # фильтр по категории
            if current_cat != '(All categories)' and cat != current_cat:
                continue

            # фильтр по поиску (имя + description)
            if query:
                name_l = name.lower()
                desc_l = desc.lower()
                if query not in name_l and query not in desc_l:
                    continue

            # ------------------ UI элемент строки ------------------
            stack = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal)
            stack.Margin = Thickness(2)

            # --- картинка ---
            img = Image()
            img.Width = 96      # 1.5x 64
            img.Height = 96
            img.Margin = Thickness(0, 0, 10, 0)

            if fam.get('preview') and os.path.exists(fam['preview']):
                try:
                    file_uri = 'file:///' + fam['preview'].replace('\\', '/')
                    bmp = BitmapImage()
                    bmp.BeginInit()
                    bmp.UriSource = Uri(file_uri, UriKind.Absolute)
                    bmp.CacheOption = BitmapCacheOption.OnLoad
                    bmp.EndInit()
                    img.Source = bmp
                except Exception as ex:
                    print("Preview load error for {}: {}".format(fam['preview'], ex))

            # --- имя семейства ---
            name_text = TextBlock()
            name_text.Text = name
            name_text.VerticalAlignment = System.Windows.VerticalAlignment.Center
            name_text.Margin = Thickness(0, 0, 20, 0)
            name_text.TextWrapping = System.Windows.TextWrapping.NoWrap
            name_text.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
            name_text.Width = 250

            # --- Description справа ---
            desc_text = TextBlock()
            desc_text.Text = desc
            desc_text.VerticalAlignment = System.Windows.VerticalAlignment.Center
            desc_text.Margin = Thickness(0, 0, 0, 0)
            desc_text.TextWrapping = System.Windows.TextWrapping.NoWrap
            desc_text.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
            desc_text.Width = 350
            desc_text.TextAlignment = System.Windows.TextAlignment.Left

            stack.Children.Add(img)
            stack.Children.Add(name_text)
            stack.Children.Add(desc_text)

            item = ListBoxItem()
            item.Content = stack
            item.Tag = fam  # словарь с данными о семействе

            self.familyList.Items.Add(item)

    # ------------------------------ XAML handlers ----------------------

    def categoryCombo_SelectionChanged(self, sender, args):
        self._populate_family_list()

    def searchBox_TextChanged(self, sender, args):
        self._populate_family_list()

    def close_button_click(self, sender, args):
        self.Close()

    def refresh_button_click(self, sender, args):
        res = forms.alert(
            u"Rebuild family library index?\n"
            u"This may take a while.",
            yes=True, no=True
        )
        if not res:
            return

        new_fams = build_index(show_alert=False)
        self.families = new_fams
        self._rebuild_categories()

    def load_button_click(self, sender, args):
        # === ЭТОТ МЕТОД НУЖЕН ДЛЯ XAML ===
        item = self.familyList.SelectedItem
        if not item:
            forms.alert(u"Please select a family first.")
            return

        fam_info = item.Tag
        fam_path = fam_info['path']
        fam_name = fam_info['name']

        if not os.path.exists(fam_path):
            forms.alert(u"Family file not found:\n{}".format(fam_path), warn_icon=True)
            return

        # ---- ПРОВЕРКА: есть ли семейство в модели ----
        existing_fams = list(
            DB.FilteredElementCollector(doc)
            .OfClass(DB.Family)
            .ToElements()
        )
        is_loaded = any(f.Name == fam_name for f in existing_fams)

        if is_loaded:
            overwrite = forms.alert(
                u"Overwrite existing family types if family is already loaded?",
                yes=True, no=True, cancel=True
            )
            if overwrite is None:
                return
            overwrite_flag = bool(overwrite)
        else:
            overwrite_flag = False   # семейства ещё нет — просто загружаем

        # ---- ЗАГРУЗКА ----
        fam_ref = StrongBox[DB.Family](None)

        t = DB.Transaction(doc, "Load approved family")
        t.Start()
        try:
            load_opts = SimpleFamilyLoadOptions(overwrite_flag)

            loaded = doc.LoadFamily(fam_path, load_opts, fam_ref)

            if loaded:
                forms.alert(u"Family '{}' loaded successfully.".format(fam_name))
            else:
                forms.alert(u"Failed to load family:\n{}".format(fam_path), warn_icon=True)

            t.Commit()
        except Exception as ex:
            t.RollBack()
            forms.alert(u"Error while loading family:\n{}\n\n{}".format(fam_path, ex), warn_icon=True)
        finally:
            self.Close()

    # ------------------------------ XAML handlers ----------------------

    def categoryCombo_SelectionChanged(self, sender, args):
        self._populate_family_list()

    def searchBox_TextChanged(self, sender, args):
        self._populate_family_list()

    def close_button_click(self, sender, args):
        self.Close()

    def refresh_button_click(self, sender, args):
        res = forms.alert(
            u"Rebuild family library index?\n"
            u"This may take a while.",
            yes=True, no=True
        )
        if not res:
            return

        new_fams = build_index(show_alert=False)
        self.families = new_fams
        self._rebuild_categories()

def load_button_click(self, sender, args):
    item = self.familyList.SelectedItem
    if not item:
        forms.alert(u"Please select a family first.")
        return

    fam_info = item.Tag
    fam_path = fam_info['path']
    fam_name = fam_info['name']

    if not os.path.exists(fam_path):
        forms.alert(u"Family file not found:\n{}".format(fam_path), warn_icon=True)
        return

    # ---- ПРОВЕРЯЕМ, ЕСТЬ ЛИ РЕАЛЬНО СЕМЕЙСТВО В МОДЕЛИ ----
    existing_fams = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.Family)
        .ToElements()
    )

    is_loaded = any(f.Name == fam_name for f in existing_fams)

    # Если семейства НЕТ в проекте → НЕ спрашиваем overwrite
    if is_loaded:
        overwrite = forms.alert(
            u"Overwrite existing family types if family is already loaded?",
            yes=True, no=True, cancel=True
        )

        if overwrite is None:
            return

        overwrite_flag = bool(overwrite)
    else:
        overwrite_flag = False   # загружаем без overwrite

    # ---- ЗАГРУЗКА ----
    from System.Runtime.CompilerServices import StrongBox
    fam_ref = StrongBox[DB.Family](None)

    t = DB.Transaction(doc, "Load approved family")
    t.Start()
    try:
        load_opts = SimpleFamilyLoadOptions(overwrite_flag)

        loaded = doc.LoadFamily(fam_path, load_opts, fam_ref)

        if loaded:
            forms.alert(u"Family '{}' loaded successfully.".format(fam_name))
        else:
            forms.alert(u"Failed to load family:\n{}".format(fam_path), warn_icon=True)

        t.Commit()
    except Exception as ex:
        t.RollBack()
        forms.alert(u"Error while loading family:\n{}\n\n{}".format(fam_path, ex), warn_icon=True)
    finally:
        self.Close()



# ======================================================================
# MAIN
# ======================================================================

families = load_index()
if families is None:
    families = build_index(show_alert=True)

window = FamilyBrowserWindow(families)
window.ShowDialog()
# ======================================================================