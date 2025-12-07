# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

# --- HELPER FUNCTIONS ---
def get_param_val(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p and p.HasValue: return p.AsString()
    except: pass
    return None

def get_hosting_info(family_symbol):
    try:
        if hasattr(family_symbol, "Family") and family_symbol.Family:
            fam = family_symbol.Family
            p_type = fam.FamilyPlacementType
            if p_type == DB.FamilyPlacementType.OneLevelBased: return "Level Based"
            elif p_type == DB.FamilyPlacementType.FaceBased: return "Face Based"
            elif p_type == DB.FamilyPlacementType.WorkPlaneBased: return "Work Plane"
            else: return "Other"
        return "?"
    except: return "?"

# ==========================================
# 1. SELECT LINK
# ==========================================
links_collector = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
loaded_links = [l for l in links_collector if DB.RevitLinkType.IsLoaded(doc, l.GetTypeId())]
if not loaded_links: forms.alert("No loaded links found.", exitscript=True)

link_dict = {l.Name: l for l in loaded_links}
selected_link_name = forms.SelectFromList.show(
    sorted(link_dict.keys()), 
    title="1. Select Linked Model", 
    button_name="Select"
)
if not selected_link_name: script.exit()

link_instance = link_dict[selected_link_name]
link_doc = link_instance.GetLinkDocument()
link_transform = link_instance.GetTotalTransform()

# ==========================================
# 2. SELECT CATEGORY (SOURCE)
# ==========================================
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
    count = DB.FilteredElementCollector(link_doc).OfCategory(bic).WhereElementIsNotElementType().GetElementCount()
    if count > 0:
        try:
            c_name = DB.Category.GetCategory(link_doc, bic).Name
            options_cats[c_name] = bic
        except: pass

if not options_cats: forms.alert("No elements found in the link.", exitscript=True)
selected_cat_name = forms.SelectFromList.show(
    sorted(options_cats.keys()), 
    title="2. Select Category (Source)", 
    button_name="Next"
)
if not selected_cat_name: script.exit()
selected_bic = options_cats[selected_cat_name]

# ==========================================
# 3. SELECT FAMILY (SOURCE)
# ==========================================
elements_in_link = DB.FilteredElementCollector(link_doc).OfCategory(selected_bic).WhereElementIsNotElementType().ToElements()
src_types = {}

for el in elements_in_link:
    tid = el.GetTypeId()
    if tid == DB.ElementId.InvalidElementId: continue
    sym = link_doc.GetElement(tid)
    if not sym or not isinstance(sym, DB.FamilySymbol): continue

    f_name = get_param_val(sym, DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    if not f_name: 
        try: f_name = sym.FamilyName
        except: f_name = "Unknown"
            
    t_name = get_param_val(sym, DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    if not t_name: 
        try: t_name = sym.Name
        except: t_name = "Unknown"
    
    dict_key = "{}: {}".format(f_name, t_name)
    if dict_key not in src_types:
        host_info = get_hosting_info(sym)
        display_name = "[{}] {}".format(host_info, dict_key)
        src_types[display_name] = {"symbol": sym, "key_name": dict_key}

if not src_types: forms.alert("No families found.", exitscript=True)

selected_src_display_name = forms.SelectFromList.show(
    sorted(src_types.keys()), 
    title="3. Select Family (Source)", 
    width=600, 
    button_name="Next"
)
if not selected_src_display_name: script.exit()
selected_src_key = src_types[selected_src_display_name]["key_name"]

# Collect source elements
target_src_elements = []
for el in elements_in_link:
    tid = el.GetTypeId()
    if tid == DB.ElementId.InvalidElementId: continue
    sym = link_doc.GetElement(tid)
    if not isinstance(sym, DB.FamilySymbol): continue
    
    f_n = get_param_val(sym, DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    if not f_n:
        try: f_n = sym.FamilyName
        except: f_n = "Unknown"
            
    t_n = get_param_val(sym, DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    if not t_n:
        try: t_n = sym.Name
        except: t_n = "Unknown"
    
    if "{}: {}".format(f_n, t_n) == selected_src_key:
        target_src_elements.append(el)

# ==========================================
# 4. MANUAL LEVEL MAPPING
# ==========================================
levels_in_link_names = set()
levels_in_link_map = {} # {LevelName: LevelObj}

for el in target_src_elements:
    try:
        lid = el.LevelId
        if lid != DB.ElementId.InvalidElementId:
            lev = link_doc.GetElement(lid)
            if lev: 
                levels_in_link_names.add(lev.Name)
                levels_in_link_map[lev.Name] = lev
    except: pass

if not levels_in_link_map:
    forms.alert("Source elements have no levels defined. Mapping impossible.", exitscript=True)

# 4.1 Select SOURCE Levels
selected_arch_levels_names = forms.SelectFromList.show(
    sorted(levels_in_link_map.keys()), 
    title="4.1 Select SOURCE Levels (from Link)", 
    multiselect=True, 
    button_name="Next"
)
if not selected_arch_levels_names: script.exit()

# 4.2 Get TARGET Levels
my_levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
my_levels_dict = {l.Name: l for l in my_levels}
my_levels_names_sorted = sorted(my_levels_dict.keys())

# 4.3 Mapping Loop
level_mapping = {} # {Arch_Level_Name : My_Level_Object}

for arch_lvl_name in selected_arch_levels_names:
    preselect = [arch_lvl_name] if arch_lvl_name in my_levels_dict else None
    
    chosen_my_lvl_name = forms.SelectFromList.show(
        my_levels_names_sorted,
        title="Link Level: '{}' -> Target Level: ???".format(arch_lvl_name),
        default=preselect,
        multiselect=False,
        button_name="Map Level"
    )
    
    if not chosen_my_lvl_name: continue
    level_mapping[arch_lvl_name] = my_levels_dict[chosen_my_lvl_name]

if not level_mapping:
    forms.alert("No level mapping defined.", exitscript=True)


# ==========================================
# 5. SELECT CATEGORY (TARGET)
# ==========================================
my_cats = {}
all_loaded_symbols = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements()
for sym in all_loaded_symbols:
    if sym.Category: my_cats[sym.Category.Name] = sym.Category.Id

preselect_cat = None
try:
    src_cat_name = DB.Category.GetCategory(link_doc, selected_bic).Name
    if src_cat_name in my_cats: preselect_cat = src_cat_name
except: pass

selected_dest_cat_name = forms.SelectFromList.show(
    sorted(my_cats.keys()), 
    title="5. Select Category (Target)", 
    default=[preselect_cat] if preselect_cat else None, 
    button_name="Select"
)
if not selected_dest_cat_name: script.exit()
dest_cat_id = my_cats[selected_dest_cat_name]

# ==========================================
# 6. SELECT FAMILY (TARGET)
# ==========================================
my_symbols = DB.FilteredElementCollector(doc).OfCategoryId(dest_cat_id).WhereElementIsElementType().ToElements()
my_types_dict = {}
for sym in my_symbols:
    if not isinstance(sym, DB.FamilySymbol): continue
    
    f_name = get_param_val(sym, DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    if not f_name:
        try: f_name = sym.FamilyName
        except: f_name = "Fam"
            
    t_name = get_param_val(sym, DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    if not t_name:
        try: t_name = sym.Name
        except: t_name = "Type"
    
    h_info = get_hosting_info(sym)
    d_name = "[{}] {}: {}".format(h_info, f_name, t_name)
    if d_name in my_types_dict: d_name = "{} (ID {})".format(d_name, sym.Id)
    my_types_dict[d_name] = sym

selected_dest_name = forms.SelectFromList.show(
    sorted(my_types_dict.keys()), 
    title="6. Select Family (Target)", 
    width=600, 
    button_name="Place Families"
)
if not selected_dest_name: script.exit()
dest_symbol = my_types_dict[selected_dest_name]

# ==========================================
# 7. EXECUTION (MAPPING + HEIGHT CORRECTION)
# ==========================================
with DB.Transaction(doc, "Activate") as t:
    t.Start()
    if not dest_symbol.IsActive: dest_symbol.Activate(); doc.Regenerate()
    t.Commit()

count_placed = 0
created_ids = List[DB.ElementId]()
log_levels = set()
errors = []

with DB.Transaction(doc, "SHN: Copy Families") as t:
    t.Start()
    for el in target_src_elements:
        try:
            # 1. Check Source Level
            if el.LevelId == DB.ElementId.InvalidElementId: continue
            
            lev_src_obj = link_doc.GetElement(el.LevelId)
            lev_src_name = lev_src_obj.Name
            
            # 2. Check Mapping
            if lev_src_name not in level_mapping: continue
            tgt_lvl = level_mapping[lev_src_name]
            
            # 3. Geometry
            loc = el.Location
            xyz_src = None
            rot = 0
            if isinstance(loc, DB.LocationPoint):
                xyz_src = loc.Point
                rot = loc.Rotation
            elif isinstance(loc, DB.LocationCurve):
                xyz_src = loc.Curve.Evaluate(0.5, True)
            else: continue
            
            xyz_target = link_transform.OfPoint(xyz_src)
            
            # 4. Create Element
            new_inst = doc.Create.NewFamilyInstance(xyz_target, dest_symbol, tgt_lvl, DB.Structure.StructuralType.NonStructural)
            doc.Regenerate()
            
            # 5. Height Correction
            if isinstance(new_inst.Location, DB.LocationPoint):
                current_z = new_inst.Location.Point.Z
            else:
                current_z = tgt_lvl.Elevation
            
            diff_z = xyz_target.Z - current_z
            
            if abs(diff_z) > 0.003: # >3mm
                move_vec = DB.XYZ(0, 0, diff_z)
                try:
                    DB.ElementTransformUtils.MoveElement(doc, new_inst.Id, move_vec)
                except: pass

            # 6. Rotation
            try:
                axis = DB.Line.CreateBound(xyz_target, xyz_target + DB.XYZ(0,0,1))
                DB.ElementTransformUtils.RotateElement(doc, new_inst.Id, axis, rot)
            except: pass
            
            count_placed += 1
            log_levels.add(tgt_lvl.Name)
            created_ids.Add(new_inst.Id)
            
        except Exception as e:
            errors.append(str(e))

    t.Commit()

# ==========================================
# 8. FINAL REPORT
# ==========================================
if count_placed > 0:
    try: uidoc.Selection.SetElementIds(created_ids)
    except: pass
    
    msg = "Success!\nCreated: {}\nLevels: {}".format(count_placed, ", ".join(list(log_levels)))
    forms.alert(msg)
else:
    msg = "Nothing created."
    if len(errors) > 0: msg += "\nError: " + errors[0]
    forms.alert(msg, warn_icon=True)