# -*- coding: utf-8 -*-
"""
Copy level-based lighting from linked model
into host model using level-based target family.
"""

__title__ = 'Family\nTransfer'
__doc__ = 'Copy level-based lighting from linked model into host model using level-based target family.'
__author__ = 'SHNABEL digital'


from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc


# ----------------------------------------------------------
# Helpers
# ----------------------------------------------------------
def get_param_val(elem, bip):
    """Безопасно прочитать строковый параметр по BuiltInParameter."""
    try:
        p = elem.get_Parameter(bip)
        if p and p.HasValue:
            return p.AsString()
    except:
        pass
    return None


def get_hosting_info(sym):
    """Информативная строка о типе размещения семейства."""
    try:
        fam = sym.Family
        if fam:
            return str(fam.FamilyPlacementType)
        return "?"
    except:
        return "?"


# ----------------------------------------------------------
# 1. Выбор линка
# ----------------------------------------------------------
links_collector = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
loaded_links = [l for l in links_collector if DB.RevitLinkType.IsLoaded(doc, l.GetTypeId())]

if not loaded_links:
    forms.alert("No loaded Revit links found in the project.", exitscript=True)

link_dict = {l.Name: l for l in loaded_links}

selected_link_name = forms.SelectFromList.show(
    sorted(link_dict.keys()),
    title="1. Select Linked Model",
    button_name="Select"
)
if not selected_link_name:
    script.exit()

link_instance = link_dict[selected_link_name]
link_doc = link_instance.GetLinkDocument()
link_transform = link_instance.GetTotalTransform()

if not link_doc:
    forms.alert("Selected link has no accessible document.", exitscript=True)


# ----------------------------------------------------------
# 2. Выбор категории в линку (источник)
# ----------------------------------------------------------
cats_to_check = [
    DB.BuiltInCategory.OST_LightingFixtures,
    DB.BuiltInCategory.OST_ElectricalEquipment,
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_CommunicationDevices,
    DB.BuiltInCategory.OST_DataDevices,
    DB.BuiltInCategory.OST_SecurityDevices,
    DB.BuiltInCategory.OST_FireAlarmDevices,
    DB.BuiltInCategory.OST_GenericModel,
    DB.BuiltInCategory.OST_SpecialityEquipment
]

options_cats = {}
for bic in cats_to_check:
    try:
        count = (DB.FilteredElementCollector(link_doc)
                 .OfCategory(bic)
                 .WhereElementIsNotElementType()
                 .GetElementCount())
    except:
        count = 0
    if count > 0:
        try:
            c_name = DB.Category.GetCategory(link_doc, bic).Name
            options_cats[c_name] = bic
        except:
            pass

if not options_cats:
    forms.alert("No source elements found in the selected link.", exitscript=True)

selected_cat_name = forms.SelectFromList.show(
    sorted(options_cats.keys()),
    title="2. Select Category (Source in Link)",
    button_name="Next"
)
if not selected_cat_name:
    script.exit()

selected_bic = options_cats[selected_cat_name]


# ----------------------------------------------------------
# 3. Выбор семейства/типа в линку (источник)
# ----------------------------------------------------------
elements_in_link = (DB.FilteredElementCollector(link_doc)
                    .OfCategory(selected_bic)
                    .WhereElementIsNotElementType()
                    .ToElements())

src_types = {}

for el in elements_in_link:
    tid = el.GetTypeId()
    if tid == DB.ElementId.InvalidElementId:
        continue
    sym = link_doc.GetElement(tid)
    if not sym or not isinstance(sym, DB.FamilySymbol):
        continue

    f_name = get_param_val(sym, DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    if not f_name:
        try:
            f_name = sym.FamilyName
        except:
            f_name = "Unknown"

    t_name = get_param_val(sym, DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    if not t_name:
        try:
            t_name = sym.Name
        except:
            t_name = "Unknown"

    dict_key = "{}: {}".format(f_name, t_name)
    if dict_key not in src_types:
        host_info = get_hosting_info(sym)
        display_name = "[{}] {}".format(host_info, dict_key)
        src_types[display_name] = {"symbol": sym, "key_name": dict_key}

if not src_types:
    forms.alert("No family types found in the selected category of the link.", exitscript=True)

selected_src_display_name = forms.SelectFromList.show(
    sorted(src_types.keys()),
    title="3. Select Family Type (Source in Link)",
    width=600,
    button_name="Next"
)
if not selected_src_display_name:
    script.exit()

selected_src_key = src_types[selected_src_display_name]["key_name"]

# Собираем только экземпляры выбранного типа
target_src_elements = []
for el in elements_in_link:
    tid = el.GetTypeId()
    if tid == DB.ElementId.InvalidElementId:
        continue
    sym = link_doc.GetElement(tid)
    if not isinstance(sym, DB.FamilySymbol):
        continue

    f_n = get_param_val(sym, DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    if not f_n:
        try:
            f_n = sym.FamilyName
        except:
            f_n = "Unknown"

    t_n = get_param_val(sym, DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    if not t_n:
        try:
            t_n = sym.Name
        except:
            t_n = "Unknown"

    if "{}: {}".format(f_n, t_n) == selected_src_key:
        target_src_elements.append(el)

if not target_src_elements:
    forms.alert("No instances of selected family type were found in the link.", exitscript=True)


# ----------------------------------------------------------
# 4. Сбор уровней в линку и маппинг на уровни в нашей модели
# ----------------------------------------------------------
levels_in_link_map = {}     # { LevelName : LevelObj }

for el in target_src_elements:
    try:
        lid = el.LevelId
        if lid != DB.ElementId.InvalidElementId:
            lev = link_doc.GetElement(lid)
            if isinstance(lev, DB.Level):
                levels_in_link_map[lev.Name] = lev
    except:
        pass

if not levels_in_link_map:
    forms.alert("Source elements have no valid levels. Mapping impossible.", exitscript=True)

# 4.1 Выбор уровней-источников, которые хотим переносить
selected_arch_levels_names = forms.SelectFromList.show(
    sorted(levels_in_link_map.keys()),
    title="4.1 Select SOURCE Levels (from Link)",
    multiselect=True,
    button_name="Next"
)
if not selected_arch_levels_names:
    script.exit()

# 4.2 Уровни в нашей модели
my_levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
my_levels_dict = {l.Name: l for l in my_levels}
my_levels_names_sorted = sorted(my_levels_dict.keys())

# 4.3 Маппинг уровней по имени
level_mapping = {}   # { Arch_Level_Name : My_Level_Object }

for arch_lvl_name in selected_arch_levels_names:
    preselect = [arch_lvl_name] if arch_lvl_name in my_levels_dict else None

    chosen_my_lvl_name = forms.SelectFromList.show(
        my_levels_names_sorted,
        title="Link Level: '{}' -> Target Level: ???".format(arch_lvl_name),
        default=preselect,
        multiselect=False,
        button_name="Map Level"
    )

    if not chosen_my_lvl_name:
        continue
    level_mapping[arch_lvl_name] = my_levels_dict[chosen_my_lvl_name]

if not level_mapping:
    forms.alert("No level mapping defined. Nothing to place.", exitscript=True)


# ----------------------------------------------------------
# 5. Выбор категории и типа в нашей модели (цель, LEVEL-BASED)
# ----------------------------------------------------------
my_cats = {}
all_loaded_symbols = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements()
for sym in all_loaded_symbols:
    try:
        if sym.Category:
            my_cats[sym.Category.Name] = sym.Category.Id
    except:
        pass

if not my_cats:
    forms.alert("No loaded family symbols found in the host model.", exitscript=True)

preselect_cat = None
try:
    src_cat_name = DB.Category.GetCategory(link_doc, selected_bic).Name
    if src_cat_name in my_cats:
        preselect_cat = src_cat_name
except:
    pass

selected_dest_cat_name = forms.SelectFromList.show(
    sorted(my_cats.keys()),
    title="5. Select Category (Target in Host Model)",
    default=[preselect_cat] if preselect_cat else None,
    button_name="Select"
)
if not selected_dest_cat_name:
    script.exit()

dest_cat_id = my_cats[selected_dest_cat_name]

my_symbols = (DB.FilteredElementCollector(doc)
              .OfCategoryId(dest_cat_id)
              .WhereElementIsElementType()
              .ToElements())

my_types_dict = {}
for sym in my_symbols:
    if not isinstance(sym, DB.FamilySymbol):
        continue

    f_name = get_param_val(sym, DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    if not f_name:
        try:
            f_name = sym.FamilyName
        except:
            f_name = "Fam"

    t_name = get_param_val(sym, DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    if not t_name:
        try:
            t_name = sym.Name
        except:
            t_name = "Type"

    h_info = get_hosting_info(sym)
    d_name = "[{}] {}: {}".format(h_info, f_name, t_name)
    if d_name in my_types_dict:
        d_name = "{} (ID {})".format(d_name, sym.Id)
    my_types_dict[d_name] = sym

if not my_types_dict:
    forms.alert("No family types found in the selected target category.", exitscript=True)

selected_dest_name = forms.SelectFromList.show(
    sorted(my_types_dict.keys()),
    title="6. Select Family (TARGET Level-Based in Host Model)",
    width=600,
    button_name="Place Families"
)
if not selected_dest_name:
    script.exit()

dest_symbol = my_types_dict[selected_dest_name]


# ----------------------------------------------------------
# 6. Размещение с коррекцией высоты
# ----------------------------------------------------------
# Активируем тип
with DB.Transaction(doc, "Activate target family type") as t:
    t.Start()
    if not dest_symbol.IsActive:
        dest_symbol.Activate()
        doc.Regenerate()
    t.Commit()

count_placed = 0
created_ids = List[DB.ElementId]()
errors = []
used_levels = set()

with DB.Transaction(doc, "SHN: Copy Level-Based Families From Link") as t:
    t.Start()

    for el in target_src_elements:
        try:
            fi = el if isinstance(el, DB.FamilyInstance) else None
            if fi is None:
                continue

            # Уровень в линку
            lid = fi.LevelId
            if lid == DB.ElementId.InvalidElementId:
                continue
            lev_src = link_doc.GetElement(lid)
            if not isinstance(lev_src, DB.Level):
                continue
            arch_lvl_name = lev_src.Name

            # Проверяем, есть ли маппинг
            if arch_lvl_name not in level_mapping:
                continue
            tgt_lvl = level_mapping[arch_lvl_name]

            # Геометрия (точка и поворот)
            loc = fi.Location
            if not isinstance(loc, DB.LocationPoint):
                continue

            pt_link = loc.Point
            rot = loc.Rotation

            # Точка в координатах нашей модели (XYZ)
            pt_host = link_transform.OfPoint(pt_link)

            # Создаём экземпляр level-based семейства (Z здесь Revit почти игнорирует)
            new_inst = doc.Create.NewFamilyInstance(
                pt_host,
                dest_symbol,
                tgt_lvl,
                DB.Structure.StructuralType.NonStructural
            )

            # Поворот вокруг вертикальной оси
            try:
                axis = DB.Line.CreateBound(
                    pt_host,
                    pt_host + DB.XYZ(0, 0, 1)
                )
                DB.ElementTransformUtils.RotateElement(doc, new_inst.Id, axis, rot)
            except:
                pass

            # ---- Коррекция высоты (Z) ----
            try:
                doc.Regenerate()
                loc_new = new_inst.Location
                if isinstance(loc_new, DB.LocationPoint):
                    current_z = loc_new.Point.Z
                    # pt_host.Z - истинная высота по линку
                    diff_z = pt_host.Z - current_z
                    if abs(diff_z) > 0.001:   # ~0.3 мм
                        move_vec = DB.XYZ(0, 0, diff_z)
                        DB.ElementTransformUtils.MoveElement(doc, new_inst.Id, move_vec)
            except:
                pass
            # -------------------------------

            count_placed += 1
            created_ids.Add(new_inst.Id)
            used_levels.add(tgt_lvl.Name)

        except Exception as e:
            errors.append(str(e))

    t.Commit()


# ----------------------------------------------------------
# 7. Финальный отчёт
# ----------------------------------------------------------
if count_placed > 0:
    try:
        uidoc.Selection.SetElementIds(created_ids)
    except:
        pass

    msg = (
        "Success!\n"
        "Link: {}\n"
        "Source family: {}\n"
        "Target family: {}\n"
        "Created instances: {}\n"
        "Levels used: {}"
    ).format(
        selected_link_name,
        selected_src_key,
        selected_dest_name,
        count_placed,
        ", ".join(sorted(list(used_levels)))
    )
    forms.alert(msg)
else:
    msg = "Nothing was created."
    if errors:
        msg += "\n\nFirst error:\n{}".format(errors[0])
    forms.alert(msg, warn_icon=True)
