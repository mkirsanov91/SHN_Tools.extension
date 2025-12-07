# -*- coding: utf-8 -*-
"""
Clean Views
"""
__title__ = 'Clean Views\nand Templates'
__doc__ = 'Hide unnecessary categories, worksets and elements from linked models according to SHN drafting standards. Use this tool to clean views and sheets before exporting or issuing drawings.'
__author__ = 'SHNABEL digital'

import clr
import System
import sys
from System import Type

from pyrevit import forms

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import Autodesk.Revit.DB as DB

doc = __revit__.ActiveUIDocument.Document

# ============================================================================
# 0. UI CONFIRMATION
# ============================================================================
message_text = (
    "Are you sure you want to clean all 'SHN_' View Templates?\n\n"
    "This action will:\n"
    "1. Reset all RVT Link overrides (remove existing settings).\n"
    "2. Set Links to 'Custom' mode.\n"
    "3. Hide specific Annotation Categories (Grids, Levels, Dims, etc.).\n\n"
    "Proceed?"
)

result = forms.alert(message_text, title="Confirm Cleaning", warn_icon=False, yes=True, no=True)

if not result:
    print("Cancelled by user.")
    sys.exit()

# ============================================================================
# 1. COMMON DATA
# ============================================================================
# Категории, которые нужно скрыть (host + links)
categories_to_hide = [
    DB.BuiltInCategory.OST_Dimensions,       # Dimensions
    DB.BuiltInCategory.OST_VolumeOfInterest, # Scope Boxes
    DB.BuiltInCategory.OST_CLines,           # Reference Planes
    DB.BuiltInCategory.OST_ReferenceLines,   # Reference Lines
    DB.BuiltInCategory.OST_SpotSlopes,       # Spot Slopes
    DB.BuiltInCategory.OST_Grids,            # Grids
    DB.BuiltInCategory.OST_Levels            # Levels
]


def enable_include_checkbox(view_template, param_id_int):
    """Включает галочку 'Include' у параметра шаблона (если она отключена)."""
    try:
        pid = DB.ElementId(param_id_int)
        non_controlled = list(view_template.GetNonControlledTemplateParameterIds())
        if pid in non_controlled:
            non_controlled.remove(pid)
            view_template.SetNonControlledTemplateParameterIds(non_controlled)
        return True
    except Exception as e:
        print("  [WARN] Include checkbox not updated (param {}): {}".format(param_id_int, e))
        return False


def set_custom_visibility(settings):
    """Устанавливает LinkVisibility = Custom стандартным API, без reflection."""
    try:
        settings.SetLinkVisibility(DB.RevitLinkVisibility.Custom)
        return True
    except Exception as e:
        print("  [WARN] Can not set Custom visibility: {}".format(e))
        return False


# Собираем линк-инстансы и шаблоны ОДИН РАЗ
all_link_instances = list(DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance))
all_views = list(DB.FilteredElementCollector(doc).OfClass(DB.View))
shn_templates = [v for v in all_views if v.IsTemplate and v.Name.startswith("SHN_")]

if not shn_templates:
    print("No view templates starting with 'SHN_' were found. Nothing to clean.")
    sys.exit()

print("Found {} 'SHN_' view templates.".format(len(shn_templates)))
print("Found {} Revit links.".format(len(all_link_instances)))

# ============================================================================
# 2. TRANSACTIONS
# ============================================================================

tgroup = DB.TransactionGroup(doc, "SHN_Clean_ViewTemplates")
tgroup.Start()

step1_ok = True

# ==========================================
# STEP 1: REMOVE OLD LINK OVERRIDES
# ==========================================
t1 = DB.Transaction(doc, "SHN_Step1_RemoveLinkOverrides")
t1.Start()
print("=== STEP 1: REMOVING OLD SETTINGS ===")

try:
    for vt in shn_templates:
        # Включаем галочку "RVT Links" (magic ID — твоя логика, оставляю)
        enable_include_checkbox(vt, -1006967)

        for link in all_link_instances:
            try:
                # Полный сброс настроек для линка
                # Вариант 1: RemoveLinkOverrides, если метод доступен в твоей версии Revit
                vt.RemoveLinkOverrides(link.Id)

                # Вариант 2 (более универсальный, если выше не работает):
                # default_settings = DB.RevitLinkGraphicsSettings()
                # vt.SetLinkOverrides(link.Id, default_settings)

            except Exception as e_inner:
                print("  [WARN] Cannot remove overrides for link '{}': {}".format(link.Name, e_inner))

    t1.Commit()
    print("Old settings removed successfully.\n")

except Exception as e:
    t1.RollBack()
    step1_ok = False
    print("Error during removal step: {}".format(e))

# Если первый шаг провалился — не продолжаем
if not step1_ok:
    tgroup.RollBack()
    sys.exit()

# ==========================================
# STEP 2: APPLY NEW CUSTOM SETTINGS
# ==========================================
t2 = DB.Transaction(doc, "SHN_Step2_ApplyCustom")
t2.Start()
print("=== STEP 2: APPLYING CUSTOM SETTINGS ===")

try:
    for vt in shn_templates:
        print("--- Template: {} ---".format(vt.Name))

        # 1. Host Annotations (размеры, оси и т.п. в нашем документе)
        # magic ID -1006964 — как у тебя, включает управление аннотациями в шаблоне
        enable_include_checkbox(vt, -1006964)

        for cat_enum in categories_to_hide:
            cat_id = DB.ElementId(cat_enum)
            try:
                if vt.CanCategoryBeHidden(cat_id):
                    if not vt.GetCategoryHidden(cat_id):
                        vt.SetCategoryHidden(cat_id, True)
            except Exception as e_cat:
                print("  [WARN] Host category hide failed for {}: {}".format(cat_enum, e_cat))

        # 2. RVT Links — ставим Custom + скрываем те же категории внутри линов
        for link in all_link_instances:
            try:
                final_settings = DB.RevitLinkGraphicsSettings()

                # Ставим Custom нормальным API
                if set_custom_visibility(final_settings):

                    # Скрываем нужные категории в линке
                    for cat_enum in categories_to_hide:
                        cat_id = DB.ElementId(cat_enum)
                        try:
                            final_settings.SetCategoryVisibility(cat_id, False)
                        except Exception:
                            # Не все категории могут быть валидны для конкретного линка — это нормально
                            pass

                    # Применяем к шаблону
                    vt.SetLinkOverrides(link.Id, final_settings)

                    # Проверка
                    check = vt.GetLinkOverrides(link.Id)
                    # LinkVisibilityType — enum RevitLinkVisibility
                    try:
                        vis_type = check.LinkVisibilityType
                        if vis_type == DB.RevitLinkVisibility.Custom:
                            print("  [OK] {} -> Custom".format(link.Name))
                        else:
                            print("  [FAIL] {} -> visibility is {}".format(link.Name, vis_type))
                    except Exception:
                        print("  [INFO] {} -> overrides applied (no visibility check)".format(link.Name))
                else:
                    print("  [FAIL] {} -> Custom not set".format(link.Name))

            except Exception as e:
                print("  Link Error {}: {}".format(link.Name, e))

    t2.Commit()
    tgroup.Assimilate()
    print("\nDone! All steps completed.")

except Exception as e:
    t2.RollBack()
    tgroup.RollBack()
    import traceback
    print("Error during apply step:")
    print(traceback.format_exc())
