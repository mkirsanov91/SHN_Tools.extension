# -*- coding: utf-8 -*-
"""
Auto-fill 'Description' parameter for elements by category and family.
Asks which source parameters to use for each family.
"""

__title__ = 'Fill\nDescription'
__author__ = 'SHNABEL Dept (for Misha)'

from pyrevit import revit, forms
from Autodesk.Revit import DB


# Имя параметра, в который пишем описание (можешь поменять)
DESC_PARAM_NAME = "Description"

doc = revit.doc

# Только нужные (электрические) категории
ALLOWED_BICS = [
    DB.BuiltInCategory.OST_LightingFixtures,
    DB.BuiltInCategory.OST_LightingDevices,
    DB.BuiltInCategory.OST_ElectricalEquipment,
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_CommunicationDevices,
    DB.BuiltInCategory.OST_DataDevices,
    DB.BuiltInCategory.OST_FireAlarmDevices,
    DB.BuiltInCategory.OST_SecurityDevices,
    DB.BuiltInCategory.OST_NurseCallDevices,
    DB.BuiltInCategory.OST_CableTray,
    DB.BuiltInCategory.OST_CableTrayFitting,
    DB.BuiltInCategory.OST_Conduit,
    DB.BuiltInCategory.OST_ConduitFitting,
    DB.BuiltInCategory.OST_GenericModel,
]


class CategoryItem(object):
    def __init__(self, category):
        self.category = category
    def __str__(self):
        return self.category.Name


class FamilyItem(object):
    def __init__(self, family_name, sample_element):
        self.family_name = family_name
        self.sample_element = sample_element
    def __str__(self):
        return self.family_name


def get_document_categories(document):
    cat_items = []
    for bic in ALLOWED_BICS:
        try:
            cat = document.Settings.Categories.get_Item(bic)
        except:
            cat = None
        if not cat or not cat.AllowsBoundParameters:
            continue
        collector = (DB.FilteredElementCollector(document)
                     .OfCategory(bic)
                     .WhereElementIsNotElementType())
        if collector.GetElementCount() > 0:
            cat_items.append(CategoryItem(cat))
    return cat_items


def get_sample_element_for_category(document, category):
    collector = (DB.FilteredElementCollector(document)
                 .OfCategoryId(category.Id)
                 .WhereElementIsNotElementType())
    return collector.FirstElement()


def get_family_name(element):
    fam_name = None
    try:
        symbol = element.Symbol
        if symbol and symbol.Family:
            fam_name = symbol.Family.Name
    except:
        fam_name = None
    if not fam_name:
        p = element.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        if p:
            fam_name = p.AsString()
    return fam_name


def get_families_for_category(document, category):
    families = {}
    collector = (DB.FilteredElementCollector(document)
                 .OfCategoryId(category.Id)
                 .WhereElementIsNotElementType())
    for el in collector:
        fam_name = get_family_name(el)
        if fam_name and fam_name not in families:
            families[fam_name] = el
    return families


def get_parameter_names_for_element(element):
    names = set()
    # экземпляр
    try:
        params = list(element.GetOrderedParameters())
    except Exception:
        params = list(element.Parameters)
    for p in params:
        defn = p.Definition
        if defn:
            names.add(defn.Name)

    # тип
    type_el = doc.GetElement(element.GetTypeId())
    if type_el:
        names.add("Type Name")
        try:
            t_params = list(type_el.GetOrderedParameters())
        except Exception:
            t_params = list(type_el.Parameters)
        for p in t_params:
            defn = p.Definition
            if defn:
                names.add(defn.Name)

    result = sorted(names)
    if "Type Name" in result:
        result.insert(0, result.pop(result.index("Type Name")))
    return result


def build_description_for_element(element, param_names):
    parts = []

    type_el = None
    type_id = element.GetTypeId()
    if type_id and type_id != DB.ElementId.InvalidElementId:
        type_el = doc.GetElement(type_id)

    for pname in param_names:
        if pname == "Type Name":
            val = None
            if type_el:
                try:
                    p_type_name = type_el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                    if p_type_name:
                        val = p_type_name.AsString() or p_type_name.AsValueString()
                    if not val:
                        val = getattr(type_el, "Name", None)
                except:
                    val = None
            if val:
                parts.append(val)
            continue

        p = element.LookupParameter(pname)
        if not p and type_el:
            p = type_el.LookupParameter(pname)
        if not p:
            continue

        if p.StorageType == DB.StorageType.String:
            val = p.AsString()
        else:
            val = p.AsValueString()
        if val:
            parts.append(val)

    return " | ".join(parts)


def set_description_on_element(element, desc_value):
    if not desc_value:
        return
    desc_param = element.LookupParameter(DESC_PARAM_NAME)
    if not desc_param or desc_param.StorageType != DB.StorageType.String:
        type_el = doc.GetElement(element.GetTypeId())
        if type_el:
            desc_param = type_el.LookupParameter(DESC_PARAM_NAME)
    if desc_param and desc_param.StorageType == DB.StorageType.String:
        desc_param.Set(desc_value)


def main():
    # 1. категории
    cat_items = get_document_categories(doc)
    if not cat_items:
        forms.alert("В документе нет элементов нужных категорий.", exitscript=True)

    selected_cats = forms.SelectFromList.show(
        cat_items,
        multiselect=True,
        title="Выбери категории для заполнения Description",
        button_name="OK"
    )
    if not selected_cats:
        return

    # {cat_id: {fam_name: [params]}}
    category_family_param_map = {}

    # 2. категории → семейства → параметры
    for cat_item in selected_cats:
        category = cat_item.category
        families = get_families_for_category(doc, category)
        if not families:
            continue

        family_items = [FamilyItem(fn, el) for fn, el in families.items()]

        selected_families = forms.SelectFromList.show(
            family_items,
            multiselect=True,
            title=u"Выбери семейства | Категория: {0}".format(category.Name),
            button_name="OK"
        )
        if not selected_families:
            continue

        for fam_item in selected_families:
            fam_name = fam_item.family_name
            sample_el = fam_item.sample_element

            param_names = get_parameter_names_for_element(sample_el)

            selected_params = forms.SelectFromList.show(
                param_names,
                multiselect=True,
                title=u"Description params | Cat: {0} | Family: {1}".format(
                    category.Name, fam_name
                ),
                button_name="Use selected"
            )

            if not selected_params:
                continue

            category_family_param_map.setdefault(category.Id, {})[fam_name] = list(selected_params)

    if not category_family_param_map:
        forms.alert("Не выбраны параметры ни для одного семейства.", exitscript=True)
        return

    # 3. заполнение
    t = DB.Transaction(doc, "Fill Description")
    t.Start()
    total_updated = 0

    try:
        for cat_id, family_map in category_family_param_map.items():
            collector = (DB.FilteredElementCollector(doc)
                         .OfCategoryId(cat_id)
                         .WhereElementIsNotElementType())
            for el in collector:
                fam_name = get_family_name(el)
                if not fam_name:
                    continue
                params_for_family = family_map.get(fam_name)
                if not params_for_family:
                    continue
                desc_value = build_description_for_element(el, params_for_family)
                set_description_on_element(el, desc_value)
                total_updated += 1
        t.Commit()
    except Exception as ex:
        t.RollBack()
        forms.alert("Ошибка при заполнении Description:\n{}".format(ex), exitscript=True)
        return

    forms.alert(
        "Готово.\nОбновлено элементов: {0}".format(total_updated),
        title="Fill Description"
    )


if __name__ == "__main__":
    main()
