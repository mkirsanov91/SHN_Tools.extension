# -*- coding: utf-8 -*-
"""
Clean Views
"""
__title__ = 'Clean\ntemplates'
__doc__ = 'Тотальная чистка: Удаление overrides -> Создание новых. С подтверждением.'
__author__ = 'SHN'

import clr
import System
import sys # Импортируем sys для остановки скрипта
from System import Type

# Импортируем формы PyRevit для диалогового окна
from pyrevit import forms

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document

# ============================================================================
# 0. ДИАЛОГОВОЕ ОКНО (UI CONFIRMATION)
# ============================================================================
message_text = (
    "Are you sure you want to clean all 'SHN_' View Templates?\n\n"
    "This action will:\n"
    "1. Reset all RVT Link overrides (remove existing settings).\n"
    "2. Set Links to 'Custom' mode.\n"
    "3. Hide specific Annotation Categories (Grids, Levels, Dims, etc.).\n\n"
    "Proceed?"
)

# Вызываем окно. yes=True и no=True создают кнопки подтверждения.
result = forms.alert(message_text, title="Confirm Cleaning", warn_icon=False, yes=True, no=True)

# Если результат не True (нажали No или закрыли крестиком) - выходим
if not result:
    print("Cancelled by user.")
    sys.exit()

# ============================================================================
# ДАЛЕЕ ИДЕТ ОСНОВНОЙ СКРИПТ
# ============================================================================

# --- СПИСОК КАТЕГОРИЙ (ID) ---
categories_to_hide = [
    BuiltInCategory.OST_Dimensions,       # Размеры
    BuiltInCategory.OST_VolumeOfInterest, # Scope Boxes
    BuiltInCategory.OST_CLines,           # Reference Planes
    BuiltInCategory.OST_ReferenceLines,   # Reference Lines
    BuiltInCategory.OST_SpotSlopes,       # Уклоны
    BuiltInCategory.OST_Grids,            # Оси
    BuiltInCategory.OST_Levels            # Уровни
]

def enable_include_checkbox(view_template, param_id_int):
    """Включает галочку 'Include'."""
    try:
        pid = ElementId(param_id_int)
        non_controlled = view_template.GetNonControlledTemplateParameterIds()
        if pid in non_controlled:
            non_controlled.Remove(pid)
            view_template.SetNonControlledTemplateParameterIds(non_controlled)
            return True
        return True
    except:
        return False

def set_custom_visibility_force(settings):
    """Ставит Custom (2) через Reflection."""
    try:
        prop_info = settings.GetType().GetProperty("LinkVisibilityType")
        enum_type = prop_info.PropertyType
        custom_value = System.Enum.ToObject(enum_type, 2)
        prop_info.SetValue(settings, custom_value, None)
        return True
    except:
        return False

# ==========================================
# ЭТАП 1: ПОЛНОЕ УДАЛЕНИЕ (DELETE OVERRIDES)
# ==========================================
t1 = Transaction(doc, "SHN_Step1_Remove")
t1.Start()
print("=== STEP 1: REMOVING OLD SETTINGS ===")

try:
    all_link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    views = FilteredElementCollector(doc).OfClass(View).ToElements()
    shn_templates = [v for v in views if v.IsTemplate and v.Name.startswith("SHN_")]
    
    for vt in shn_templates:
        # Включаем галочку RVT Links
        enable_include_checkbox(vt, -1006967)
        
        for link in all_link_instances:
            try:
                # ВМЕСТО ПЕРЕЗАПИСИ - УДАЛЯЕМ
                # Это полностью стирает любые данные о LinkedViewId из памяти шаблона
                vt.RemoveLinkOverrides(link.Id)
            except:
                pass 
    
    t1.Commit()
    print("Old settings removed successfully.\n")

except Exception as e:
    t1.RollBack()
    print("Error during removal step: {}".format(e))


# ==========================================
# ЭТАП 2: СОЗДАНИЕ НОВЫХ (CUSTOM)
# ==========================================
t2 = Transaction(doc, "SHN_Step2_ApplyCustom")
t2.Start()
print("=== STEP 2: APPLYING CUSTOM SETTINGS ===")

try:
    # Заново собираем шаблоны не нужно, используем тот же список shn_templates
    for vt in shn_templates:
        print("--- Template: {} ---".format(vt.Name))
        
        # 1. Main Annotations
        if enable_include_checkbox(vt, -1006964):
            pass
        for cat_enum in categories_to_hide:
            cat_id = ElementId(cat_enum)
            if vt.CanCategoryBeHidden(cat_id) and not vt.GetCategoryHidden(cat_id):
                vt.SetCategoryHidden(cat_id, True)
        
        # 2. RVT Links
        for link in all_link_instances:
            try:
                # Создаем чистейшие настройки
                final_settings = RevitLinkGraphicsSettings()
                
                # Мы НЕ трогаем LinkedViewId (по дефолту он уже Invalid)

                # Ставим Custom
                if set_custom_visibility_force(final_settings):
                    
                    # Скрываем категории
                    for cat_enum in categories_to_hide:
                        cat_id = ElementId(cat_enum)
                        try:
                            final_settings.SetCategoryVisibility(cat_id, False)
                        except:
                            pass
                    
                    # Применяем
                    vt.SetLinkOverrides(link.Id, final_settings)
                    
                    # Проверка
                    check = vt.GetLinkOverrides(link.Id)
                    val = int(check.LinkVisibilityType)
                    if val == 2:
                        print("  [OK] {} -> Custom".format(link.Name))
                    else:
                        print("  [FAIL] {} -> Not switched".format(link.Name))
                else:
                    print("  [FAIL] Reflection Error")

            except Exception as e:
                # Выводим конкретное сообщение об ошибке
                msg = e.message if hasattr(e, 'message') else str(e)
                print("  Link Error {}: {}".format(link.Name, msg))

    t2.Commit()
    print("\nDone! All steps completed.")

except Exception as e:
    t2.RollBack()
    import traceback
    print(traceback.format_exc())