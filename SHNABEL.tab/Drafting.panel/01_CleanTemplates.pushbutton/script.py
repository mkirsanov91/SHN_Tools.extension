# -*- coding: utf-8 -*-
"""
SHN Clean Views / Templates (robust)
- Clean by: Active View (or Sheet), Pick Views/Sheets, Views on Sheets (all),
            SHN_ View Templates, Views using SHN_ templates
- Reset + Apply link overrides
- Hide host categories (dims/grids/levels etc.)
- Attempt to set LinkVisibility.Custom (works in newer Revit),
  auto-detect & gracefully degrade if Custom not supported by API.
"""

__title__  = 'Clean Views\nand Templates'
__doc__    = 'Hide unnecessary categories and linked annotations according to SHN drafting standards. Choose what to clean.'
__author__ = 'SHNABEL digital'

import clr
import sys
import traceback
from pyrevit import forms

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import Autodesk.Revit.DB as DB

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
app   = __revit__.Application

# ----------------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------------
categories_to_hide = [
    DB.BuiltInCategory.OST_Dimensions,        # Dimensions
    DB.BuiltInCategory.OST_VolumeOfInterest,  # Scope Boxes
    DB.BuiltInCategory.OST_CLines,            # Reference Planes
    DB.BuiltInCategory.OST_ReferenceLines,    # Reference Lines
    DB.BuiltInCategory.OST_SpotSlopes,        # Spot Slopes
    DB.BuiltInCategory.OST_Grids,             # Grids
    DB.BuiltInCategory.OST_Levels             # Levels
]

# Magic template include ids (best-effort; may vary by version)
MAGIC_INCLUDE_RVT_LINKS    = -1006967
MAGIC_INCLUDE_ANNOTATIONS  = -1006964

# Auto-detected capability:
# None = unknown yet, True/False after first attempt
CUSTOM_SUPPORTED = None
CUSTOM_INFO_PRINTED = False


# ----------------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------------
def enable_include_checkbox(view_template, param_id_int):
    """Enable template include checkbox by removing param id from NonControlled ids (best-effort)."""
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


def is_sheet(v):
    try:
        return v and v.ViewType == DB.ViewType.DrawingSheet
    except Exception:
        return False


def is_cleanable_view(v):
    """Non-template view (sheet allowed elsewhere)."""
    try:
        if v is None:
            return False
        if hasattr(v, 'IsTemplate') and v.IsTemplate:
            return False
        return True
    except Exception:
        return False


def hide_host_categories(view_or_template, cat_enums):
    """Hide categories in host view/template (best-effort)."""
    for ce in cat_enums:
        cid = DB.ElementId(ce)
        try:
            if view_or_template.CanCategoryBeHidden(cid):
                if not view_or_template.GetCategoryHidden(cid):
                    view_or_template.SetCategoryHidden(cid, True)
        except Exception:
            pass


def collect_link_type_and_instance_ids():
    """Collect unique RevitLinkType ids and all instance ids."""
    insts = list(DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance))
    type_ids = []
    inst_ids = []
    seen_types = set()

    for li in insts:
        try:
            inst_ids.append(li.Id)
            tid = li.GetTypeId()
            if tid and tid.IntegerValue != -1 and tid.IntegerValue not in seen_types:
                seen_types.add(tid.IntegerValue)
                type_ids.append(tid)
        except Exception:
            pass

    return insts, type_ids, inst_ids


def get_views_from_sheet(sheet):
    """All views placed on a given sheet via Viewport."""
    res = []
    try:
        vps = sheet.GetAllViewports()  # ICollection<ElementId>
        for vpid in vps:
            vp = doc.GetElement(vpid)
            if not vp:
                continue
            v = doc.GetElement(vp.ViewId)
            if is_cleanable_view(v) and not is_sheet(v):
                res.append(v)
    except Exception:
        pass
    return res


def collect_views_on_all_sheets():
    """All unique views placed on any sheet (via Viewport)."""
    views = {}
    for vp in DB.FilteredElementCollector(doc).OfClass(DB.Viewport):
        try:
            v = doc.GetElement(vp.ViewId)
            if is_cleanable_view(v) and not is_sheet(v):
                views[v.Id.IntegerValue] = v
        except Exception:
            pass
    return list(views.values())


def collect_views_using_shn_templates():
    """All non-template views whose assigned template starts with SHN_."""
    all_views = list(DB.FilteredElementCollector(doc).OfClass(DB.View))
    res = []
    for v in all_views:
        if not is_cleanable_view(v):
            continue
        try:
            tid = v.ViewTemplateId
            if tid and tid.IntegerValue != -1:
                t = doc.GetElement(tid)
                if t and t.IsTemplate and t.Name.startswith("SHN_"):
                    res.append(v)
        except Exception:
            pass
    return res


def expand_targets_from_views(views):
    """
    If view has template -> target template (clean template)
    else -> target view itself.
    If input includes a SHEET -> expand to views on that sheet.
    """
    targets_views = {}
    targets_templates = {}
    mapping = []  # (viewName, templateName)

    for v in views:
        if v is None:
            continue

        # If sheet -> expand
        if is_sheet(v):
            for sv in get_views_from_sheet(v):
                try:
                    tid = sv.ViewTemplateId
                    if tid and tid.IntegerValue != -1:
                        tpl = doc.GetElement(tid)
                        if tpl and tpl.IsTemplate:
                            targets_templates[tpl.Id.IntegerValue] = tpl
                            mapping.append((sv.Name, tpl.Name))
                        else:
                            targets_views[sv.Id.IntegerValue] = sv
                    else:
                        targets_views[sv.Id.IntegerValue] = sv
                except Exception:
                    targets_views[sv.Id.IntegerValue] = sv
            continue

        # Normal view
        try:
            tid = v.ViewTemplateId
            if tid and tid.IntegerValue != -1:
                tpl = doc.GetElement(tid)
                if tpl and tpl.IsTemplate:
                    targets_templates[tpl.Id.IntegerValue] = tpl
                    mapping.append((v.Name, tpl.Name))
                else:
                    targets_views[v.Id.IntegerValue] = v
            else:
                targets_views[v.Id.IntegerValue] = v
        except Exception:
            targets_views[v.Id.IntegerValue] = v

    return list(targets_views.values()), list(targets_templates.values()), mapping


def view_supports_link_overrides(v, sample_link_type_id):
    """Detect if view/template supports link overrides (avoid spam)."""
    if v is None or sample_link_type_id is None:
        return False

    # Sheets never support link overrides (must edit views on sheets)
    if is_sheet(v):
        return False

    try:
        # Some view types typically don't support link overrides
        vt = v.ViewType
        if vt in [DB.ViewType.Schedule, DB.ViewType.Legend, DB.ViewType.Report, DB.ViewType.DraftingView]:
            return False
    except Exception:
        pass

    try:
        v.GetLinkOverrides(sample_link_type_id)
        return True
    except Exception:
        return False


# ----------------------------------------------------------------------------
# LINK GRAPHICS SETTINGS (version-robust)
# ----------------------------------------------------------------------------
def can_make_rlgs():
    return hasattr(DB, 'RevitLinkGraphicsSettings')


def try_set_visibility_custom(settings):
    """
    Correct API: settings.LinkVisibilityType = DB.LinkVisibility.Custom
    (works in some versions; may throw / be unsupported in others)
    """
    try:
        if hasattr(settings, 'LinkVisibilityType') and hasattr(DB, 'LinkVisibility'):
            settings.LinkVisibilityType = DB.LinkVisibility.Custom
            return True
    except Exception:
        pass
    return False


def try_hide_category_in_link(settings, cat_id):
    """Best-effort: settings.SetCategoryVisibility(catId, False) if available."""
    try:
        if hasattr(settings, 'SetCategoryVisibility'):
            settings.SetCategoryVisibility(cat_id, False)
            return True
    except Exception:
        pass
    return False


def apply_link_overrides(view_or_template, link_type_ids, cat_enums):
    """
    Apply link overrides to LINK TYPES (not instances) to reduce duplicates.
    Auto-detect support for Custom visibility; degrade gracefully if not supported.
    """
    global CUSTOM_SUPPORTED, CUSTOM_INFO_PRINTED

    if not can_make_rlgs():
        # no class at all -> cannot do link overrides this way
        if not CUSTOM_INFO_PRINTED:
            CUSTOM_INFO_PRINTED = True
            print("  [INFO] RevitLinkGraphicsSettings is not available in this Revit/API. "
                  "Link overrides will be skipped. Host category hiding still runs.")
        return 0, 0  # ok, warn

    ok = 0
    warn = 0

    for tid in link_type_ids:
        s = DB.RevitLinkGraphicsSettings()

        # attempt Custom only if not known false
        tried_custom = False
        if CUSTOM_SUPPORTED is not False:
            tried_custom = try_set_visibility_custom(s)

        # hide categories inside link (best-effort)
        for ce in cat_enums:
            try_hide_category_in_link(s, DB.ElementId(ce))

        try:
            view_or_template.SetLinkOverrides(tid, s)
            ok += 1

            # if we successfully set overrides after trying custom, assume supported
            if tried_custom and CUSTOM_SUPPORTED is None:
                CUSTOM_SUPPORTED = True

        except Exception as e:
            msg = str(e)

            # Typical when Custom is not supported via API in this Revit version
            # -> stop trying Custom further, and keep going (host hide still applies)
            if tried_custom and (("Custom" in msg and "not supported" in msg) or ("LinkVisibility.Custom" in msg)):
                if CUSTOM_SUPPORTED is None or CUSTOM_SUPPORTED is True:
                    CUSTOM_SUPPORTED = False
                if not CUSTOM_INFO_PRINTED:
                    CUSTOM_INFO_PRINTED = True
                    print("  [INFO] This Revit version does NOT support LinkVisibility.Custom via API.\n"
                          "        Я пропускаю настройку категорий ВНУТРИ линков и продолжаю чистить host-категории.\n"
                          "        Если нужно именно скрывать Grids/Levels/Dimensions внутри ссылок — это работает в более новых версиях Revit,\n"
                          "        либо можно использовать режим ByLinkedView и заранее подготовленные 'clean' виды в линк-файлах.")
                # do not spam; just count warning once per type
                warn += 1
                continue

            warn += 1
            print("  [WARN] Link override failed in '{}': {}".format(view_or_template.Name, e))

    return ok, warn


def reset_link_overrides(view_or_template, link_type_ids, link_inst_ids):
    """
    Reset link overrides:
    - Prefer RemoveLinkOverrides if present,
    - else SetLinkOverrides(id, default RevitLinkGraphicsSettings()) best-effort.
    """
    if not can_make_rlgs():
        return 0  # warns

    warns = 0
    has_remove = hasattr(view_or_template, 'RemoveLinkOverrides')

    for tid in link_type_ids:
        try:
            if has_remove:
                view_or_template.RemoveLinkOverrides(tid)
            else:
                view_or_template.SetLinkOverrides(tid, DB.RevitLinkGraphicsSettings())
        except Exception as e:
            warns += 1
            print("  [WARN] Reset failed (type) in '{}': {}".format(view_or_template.Name, e))

    # Instances: best-effort, usually not needed if types handled
    for iid in link_inst_ids:
        try:
            if has_remove:
                view_or_template.RemoveLinkOverrides(iid)
            else:
                view_or_template.SetLinkOverrides(iid, DB.RevitLinkGraphicsSettings())
        except Exception:
            pass

    return warns


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
mode = forms.CommandSwitchWindow.show(
    [
        "Active View (or Sheet)",
        "Pick Views / Sheets...",
        "Views on Sheets (all)",
        "SHN_ View Templates",
        "Views using SHN_ templates"
    ],
    message="Choose what to clean:"
)

if not mode:
    print("Cancelled.")
    sys.exit()

if not forms.alert(
    "This will:\n"
    "1) Reset RVT link overrides\n"
    "2) Hide host categories\n"
    "3) Try to apply link overrides (Custom + hide categories in links if API allows)\n\n"
    "Proceed?",
    title="Confirm Cleaning",
    warn_icon=False,
    yes=True,
    no=True
):
    print("Cancelled.")
    sys.exit()

print("Revit Version: {} (Build: {})".format(getattr(app, "VersionNumber", "?"), getattr(app, "VersionBuild", "?")))

# ----------------------------------------------------------------------------
# Collect targets
# ----------------------------------------------------------------------------
all_views_all = list(DB.FilteredElementCollector(doc).OfClass(DB.View))
shn_templates = [v for v in all_views_all if v.IsTemplate and v.Name.startswith("SHN_")]

targets_views = []
targets_templates = []
mapping_info = []

if mode == "Active View (or Sheet)":
    av = doc.ActiveView
    if is_sheet(av):
        views_on_sheet = get_views_from_sheet(av)
        if not views_on_sheet:
            forms.alert("Active view is a sheet but no viewports were found on it.", title="Nothing to clean", warn_icon=True)
            sys.exit()
        targets_views, targets_templates, mapping_info = expand_targets_from_views(views_on_sheet)
    else:
        targets_views, targets_templates, mapping_info = expand_targets_from_views([av])

elif mode == "Pick Views / Sheets...":
    candidates = []
    for v in all_views_all:
        # allow non-template views + sheets
        if is_sheet(v) or is_cleanable_view(v):
            candidates.append(v)

    picked = forms.SelectFromList.show(
        candidates,
        name_attr='Name',
        title="Pick Views and/or Sheets",
        multiselect=True
    )
    if not picked:
        print("No selection. Cancelled.")
        sys.exit()

    targets_views, targets_templates, mapping_info = expand_targets_from_views(picked)

elif mode == "Views on Sheets (all)":
    placed = collect_views_on_all_sheets()
    if not placed:
        forms.alert("No views found on sheets.", title="Nothing to clean", warn_icon=True)
        sys.exit()
    targets_views, targets_templates, mapping_info = expand_targets_from_views(placed)

elif mode == "SHN_ View Templates":
    if not shn_templates:
        forms.alert("No view templates starting with 'SHN_' were found.", title="Nothing to clean", warn_icon=True)
        sys.exit()
    targets_templates = shn_templates

elif mode == "Views using SHN_ templates":
    vlist = collect_views_using_shn_templates()
    if not vlist:
        forms.alert("No views found that use SHN_ templates.", title="Nothing to clean", warn_icon=True)
        sys.exit()
    targets_views, targets_templates, mapping_info = expand_targets_from_views(vlist)

targets_views = list({v.Id.IntegerValue: v for v in targets_views}.values())
targets_templates = list({t.Id.IntegerValue: t for t in targets_templates}.values())

print("\n=== MODE: {} ===".format(mode))
print("Targets: {} views (without templates), {} templates".format(len(targets_views), len(targets_templates)))

if mapping_info:
    print("\nViews controlled by templates (cleaning templates instead of views):")
    for vn, tn in mapping_info[:40]:
        print("- '{}' -> template '{}'".format(vn, tn))
    if len(mapping_info) > 40:
        print("... and {} more".format(len(mapping_info) - 40))

# ----------------------------------------------------------------------------
# Links
# ----------------------------------------------------------------------------
all_link_instances, link_type_ids, link_inst_ids = collect_link_type_and_instance_ids()
print("\nFound {} Revit link instances, {} unique link types.".format(len(all_link_instances), len(link_type_ids)))

sample_link_type_id = link_type_ids[0] if link_type_ids else None
if not sample_link_type_id:
    print("No Revit links found. Will only hide host categories.")

# ----------------------------------------------------------------------------
# Transactions
# ----------------------------------------------------------------------------
tgroup = DB.TransactionGroup(doc, "SHN_Clean_Views_Templates")
tgroup.Start()

# STEP 1: RESET
t1 = DB.Transaction(doc, "SHN_Step1_ResetLinkOverrides")
t1.Start()
print("\n=== STEP 1: RESET LINK OVERRIDES ===")

reset_warn = 0
reset_skip = 0

try:
    # Templates
    for vt in targets_templates:
        if sample_link_type_id and not view_supports_link_overrides(vt, sample_link_type_id):
            reset_skip += 1
            print("  [SKIP] '{}' does not support link overrides.".format(vt.Name))
            continue

        enable_include_checkbox(vt, MAGIC_INCLUDE_RVT_LINKS)
        if sample_link_type_id:
            reset_warn += reset_link_overrides(vt, link_type_ids, link_inst_ids)

    # Views
    for v in targets_views:
        if sample_link_type_id and not view_supports_link_overrides(v, sample_link_type_id):
            reset_skip += 1
            print("  [SKIP] '{}' does not support link overrides.".format(v.Name))
            continue

        if sample_link_type_id:
            reset_warn += reset_link_overrides(v, link_type_ids, link_inst_ids)

    t1.Commit()
    print("Step 1 done. Reset warnings: {} | Skipped: {}".format(reset_warn, reset_skip))

except Exception:
    t1.RollBack()
    tgroup.RollBack()
    print("ERROR in Step 1:\n{}".format(traceback.format_exc()))
    sys.exit()

# STEP 2: APPLY
t2 = DB.Transaction(doc, "SHN_Step2_ApplySettings")
t2.Start()
print("\n=== STEP 2: APPLY SETTINGS ===")

apply_ok = 0
apply_warn = 0
apply_skip = 0

try:
    # Templates
    for vt in targets_templates:
        print("\n--- Template: {} ---".format(vt.Name))

        # host categories always
        try:
            enable_include_checkbox(vt, MAGIC_INCLUDE_ANNOTATIONS)
            enable_include_checkbox(vt, MAGIC_INCLUDE_RVT_LINKS)
        except Exception:
            pass
        hide_host_categories(vt, categories_to_hide)

        # link overrides only if supported + links exist
        if sample_link_type_id and view_supports_link_overrides(vt, sample_link_type_id):
            ok, warn = apply_link_overrides(vt, link_type_ids, categories_to_hide)
            apply_ok += ok
            apply_warn += warn
        else:
            apply_skip += 1
            print("  [SKIP] Link overrides not supported here (or no links).")

    # Views
    for v in targets_views:
        print("\n--- View: {} ---".format(v.Name))

        hide_host_categories(v, categories_to_hide)

        if sample_link_type_id and view_supports_link_overrides(v, sample_link_type_id):
            ok, warn = apply_link_overrides(v, link_type_ids, categories_to_hide)
            apply_ok += ok
            apply_warn += warn
        else:
            apply_skip += 1
            print("  [SKIP] Link overrides not supported here (or no links).")

    t2.Commit()
    tgroup.Assimilate()

    print("\nDone!")
    print("Applied OK link-override operations: {}".format(apply_ok))
    print("Warnings: {}".format(apply_warn))
    print("Skipped targets (no link override support / no links): {}".format(apply_skip))

except Exception:
    t2.RollBack()
    tgroup.RollBack()
    print("ERROR in Step 2:\n{}".format(traceback.format_exc()))
    sys.exit()
