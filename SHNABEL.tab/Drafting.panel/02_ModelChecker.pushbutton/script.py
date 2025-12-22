# -*- coding: utf-8 -*-
from __future__ import print_function
__persistentengine__ = True
import clr
import math
import json
import os

clr.AddReference('System')
clr.AddReference('System.Core')
clr.AddReference('System.Xml')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.IO import StringReader
from System.Xml import XmlReader
from System.Collections.Generic import List
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    ElementId, Options, ViewDetailLevel, Solid, GeometryInstance,
    BooleanOperationsUtils, BooleanOperationsType, SolidUtils,
    XYZ, Transform, CurveLoop, Line, GeometryCreationUtilities,
    BoundingBoxXYZ, View3D, ViewFamily, ViewFamilyType, Transaction,
    RevitLinkInstance, FamilyInstance, StorageType, ElementCategoryFilter,
    ElementMulticategoryFilter, SetComparisonResult
)

# TemporaryViewMode exists in most versions, but keep safe
try:
    from Autodesk.Revit.DB import TemporaryViewMode
except:
    TemporaryViewMode = None

from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from pyrevit import script, forms

logger = script.get_logger()
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# ----------------------------
# Issue Status System
# ----------------------------
class IssueStatus:
    """Статусы для замечаний"""
    OPEN = "OPEN"           # Не исправлено (красный/оранжевый)
    RESOLVED = "RESOLVED"   # Исправлено (зеленый)
    IGNORED = "IGNORED"     # Не требует исправления (желтый)

class IssueStorage:
    """Хранилище статусов замечаний"""
    def __init__(self):
        self.storage_file = self.get_storage_path()
        self.data = self.load()
    
    def get_storage_path(self):
        """Путь к файлу хранения (рядом с моделью или в temp)"""
        try:
            if doc.PathName:
                # Сохраняем рядом с моделью
                model_dir = os.path.dirname(doc.PathName)
                model_name = os.path.splitext(os.path.basename(doc.PathName))[0]
                storage_dir = os.path.join(model_dir, "SHN_ModelChecker")
                if not os.path.exists(storage_dir):
                    os.makedirs(storage_dir)
                return os.path.join(storage_dir, "{0}_issues.json".format(model_name))
            else:
                # Несохраненная модель - используем temp
                import tempfile
                temp_dir = tempfile.gettempdir()
                return os.path.join(temp_dir, "SHN_ModelChecker_temp.json")
        except:
            import tempfile
            temp_dir = tempfile.gettempdir()
            return os.path.join(temp_dir, "SHN_ModelChecker_temp.json")
    
    def load(self):
        """Загрузить данные из файла"""
        try:
            if os.path.exists(self.storage_file):
                with open(self.storage_file, 'r') as f:
                    return json.load(f)
            return {}
        except:
            return {}
    
    def save(self):
        """Сохранить данные в файл"""
        try:
            with open(self.storage_file, 'w') as f:
                json.dump(self.data, f, indent=2)
            logger.info("Issue statuses saved to: {0}".format(self.storage_file))
        except Exception as ex:
            logger.error("Failed to save issue statuses: {0}".format(ex))
    
    def get_issue_key(self, issue):
        """Создать уникальный ключ для issue"""
        # Используем check_name + main_id + message
        return "{0}|{1}|{2}".format(
            issue.check_name,
            issue.main_id.IntegerValue if hasattr(issue.main_id, 'IntegerValue') else str(issue.main_id),
            issue.message[:50]  # первые 50 символов message
        )
    
    def get_status(self, issue):
        """Получить статус issue"""
        key = self.get_issue_key(issue)
        return self.data.get(key, {}).get('status', IssueStatus.OPEN)
    
    def set_status(self, issue, status):
        """Установить статус issue"""
        key = self.get_issue_key(issue)
        if key not in self.data:
            self.data[key] = {}
        self.data[key]['status'] = status
        self.data[key]['timestamp'] = str(__import__('datetime').datetime.now())
        self.save()
    
    def get_all_resolved(self):
        """Получить все resolved issues"""
        return {k: v for k, v in self.data.items() if v.get('status') == IssueStatus.RESOLVED}

# ----------------------------
# CONFIG (tune for SHN)
# ----------------------------
MM_IN_FT = 304.8

SERVICE_MM = 800.0
TRANS_CRIT_MM = 800.0
TRANS_WARN_MM = 1000.0
TRANS_HEIGHT_CHECK_MM = 1000.0  # Высота проверки ВВЕРХ над трансформатором (можно уменьшить если перекрытия мешают)
CORRIDOR_CLEARANCE_MM = 800.0
DOUBLE_ROW_SEARCH_RADIUS_MM = 8000.0

# ОТЛАДКА Check 4: Можно временно увеличить зоны чтобы понять откуда ошибки
# Например: TRANS_CRIT_MM = 600.0, TRANS_WARN_MM = 800.0
# Тогда всё что было CRITICAL станет WARNING, а WARNING станет OK

HOST_OBSTACLE_BIC = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_Conduit,
]

LINK_OBSTACLE_BIC = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_Floors,
]

TRANSFORMER_NAME_KEYS = ["transform", u"трансформ", "tr-", "tr_", "trx", "xfmr"]
SHN_CLASS_PARAM = "SHN_EquipmentClass"   # e.g. "LV", "MV", "TR"

# CHECK 4: Категории из связанных моделей для проверки трансформаторов
# Можно временно закомментировать категории для отладки
CHECK4_LINK_CATEGORIES = [
    BuiltInCategory.OST_Walls,                    # Стены
    BuiltInCategory.OST_StructuralColumns,        # Колонны
    BuiltInCategory.OST_StructuralFraming,        # Балки
    # BuiltInCategory.OST_Floors,                 # Перекрытия - ОТКЛЮЧЕНО (не блокируют доступ сбоку)
    BuiltInCategory.OST_Roofs,                    # Крыши
    BuiltInCategory.OST_StructuralFoundation,     # Фундаменты
    # BuiltInCategory.OST_Doors,                  # Двери (по умолчанию отключено)
    # BuiltInCategory.OST_Windows,                # Окна (по умолчанию отключено)
]

# ----------------------------
# Utils
# ----------------------------
def mm_to_ft(mm):
    return float(mm) / MM_IN_FT

def safe_str(x):
    try:
        return str(x)
    except:
        try:
            return x.ToString()
        except:
            return "<str?>"

def get_param_str(el, pname):
    try:
        p = el.LookupParameter(pname)
        if not p:
            return None
        if p.StorageType == StorageType.String:
            return p.AsString()
        if p.StorageType == StorageType.Integer:
            return safe_str(p.AsInteger())
        if p.StorageType == StorageType.Double:
            return safe_str(p.AsDouble())
        if p.StorageType == StorageType.ElementId:
            return safe_str(p.AsElementId().IntegerValue)
    except:
        return None
    return None

def normalize(v):
    try:
        if v is None:
            return None
        l = v.GetLength()
        if l < 1e-9:
            return None
        return XYZ(v.X/l, v.Y/l, v.Z/l)
    except:
        return None

def bbox_intersects(bb1, bb2):
    if not bb1 or not bb2:
        return False
    a_min, a_max = bb1.Min, bb1.Max
    b_min, b_max = bb2.Min, bb2.Max
    if a_max.X < b_min.X or a_min.X > b_max.X: return False
    if a_max.Y < b_min.Y or a_min.Y > b_max.Y: return False
    if a_max.Z < b_min.Z or a_min.Z > b_max.Z: return False
    return True

def bbox_expand(bb, delta_ft):
    nbb = BoundingBoxXYZ()
    nbb.Min = XYZ(bb.Min.X - delta_ft, bb.Min.Y - delta_ft, bb.Min.Z - delta_ft)
    nbb.Max = XYZ(bb.Max.X + delta_ft, bb.Max.Y + delta_ft, bb.Max.Z + delta_ft)
    return nbb

def bbox_from_points(pts):
    xs = [p.X for p in pts]; ys = [p.Y for p in pts]; zs = [p.Z for p in pts]
    bb = BoundingBoxXYZ()
    bb.Min = XYZ(min(xs), min(ys), min(zs))
    bb.Max = XYZ(max(xs), max(ys), max(zs))
    return bb

def transform_bbox(bb, trf):
    if not bb or not trf:
        return bb
    mn, mx = bb.Min, bb.Max
    pts = [
        XYZ(mn.X, mn.Y, mn.Z),
        XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mn.Z),
        XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mn.Y, mn.Z),
        XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mx.X, mx.Y, mn.Z),
        XYZ(mx.X, mx.Y, mx.Z),
    ]
    return bbox_from_points([trf.OfPoint(p) for p in pts])

def try_get_bbox(el, view=None):
    try:
        return el.get_BoundingBox(view)
    except:
        return None

def get_solids(el, geom_opt, extra_transform=None):
    solids = []
    try:
        ge = el.get_Geometry(geom_opt)
        if not ge:
            return solids
        for gobj in ge:
            if isinstance(gobj, Solid):
                if gobj.Volume and gobj.Volume > 1e-6:
                    s = gobj
                    if extra_transform:
                        try:
                            s = SolidUtils.CreateTransformed(s, extra_transform)
                        except:
                            pass
                    solids.append(s)
            elif isinstance(gobj, GeometryInstance):
                inst = gobj
                inst_tr = inst.Transform
                if extra_transform:
                    inst_tr = extra_transform.Multiply(inst_tr)
                sym_geo = inst.GetSymbolGeometry()
                if not sym_geo:
                    continue
                for sg in sym_geo:
                    if isinstance(sg, Solid) and sg.Volume and sg.Volume > 1e-6:
                        s = sg
                        try:
                            s = SolidUtils.CreateTransformed(s, inst_tr)
                        except:
                            pass
                        solids.append(s)
    except:
        pass
    return solids

def solids_intersect(solids_a, solids_b, vol_tol=1e-6):
    if not solids_a or not solids_b:
        return False
    for sa in solids_a:
        for sb in solids_b:
            try:
                inter = BooleanOperationsUtils.ExecuteBooleanOperation(sa, sb, BooleanOperationsType.Intersect)
                if inter and inter.Volume and inter.Volume > vol_tol:
                    return True
            except:
                continue
    return False

def make_local_box_solid(minX, maxX, minY, maxY, minZ, maxZ):
    """Создание solid из локальных координат (прямоугольник)"""
    p0 = XYZ(minX, minY, minZ)
    p1 = XYZ(maxX, minY, minZ)
    p2 = XYZ(maxX, maxY, minZ)
    p3 = XYZ(minX, maxY, minZ)

    cl = CurveLoop()
    cl.Append(Line.CreateBound(p0, p1))
    cl.Append(Line.CreateBound(p1, p2))
    cl.Append(Line.CreateBound(p2, p3))
    cl.Append(Line.CreateBound(p3, p0))

    loops = List[CurveLoop]()
    loops.Add(cl)

    height = maxZ - minZ
    try:
        solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, XYZ.BasisZ, height)
        return solid
    except:
        return None

def bbox_union(bb1, bb2):
    """Объединение двух bbox"""
    if not bb1:
        return bb2
    if not bb2:
        return bb1
    bb = BoundingBoxXYZ()
    bb.Min = XYZ(
        min(bb1.Min.X, bb2.Min.X),
        min(bb1.Min.Y, bb2.Min.Y),
        min(bb1.Min.Z, bb2.Min.Z)
    )
    bb.Max = XYZ(
        max(bb1.Max.X, bb2.Max.X),
        max(bb1.Max.Y, bb2.Max.Y),
        max(bb1.Max.Z, bb2.Max.Z)
    )
    return bb

# ----------------------------
# Severity
# ----------------------------
class Severity:
    CRITICAL = "CRITICAL"
    WARNING = "WARNING"
    INFO = "INFO"

# ----------------------------
# Issue class
# ----------------------------
class Issue(object):
    def __init__(self, severity, check_name, main_id, message, issue_bbox=None):
        self.severity = severity
        self.check_name = check_name
        self.main_id = main_id
        self.message = message
        self.issue_bbox = issue_bbox
        self.status = IssueStatus.OPEN  # По умолчанию OPEN
        
        # Lists of interfering elements
        self.host_blocker_ids = []
        self.link_blockers = []  # list of (linkInstanceId, linkedElementId)
    
    def add_host_blocker(self, eid):
        if eid and eid not in self.host_blocker_ids:
            self.host_blocker_ids.append(eid)
    
    def add_link_blocker(self, link_id, linked_elem_id):
        pair = (link_id, linked_elem_id)
        if pair not in self.link_blockers:
            self.link_blockers.append(pair)
    
    def interference_text(self):
        parts = []
        if self.host_blocker_ids:
            parts.append("Host: " + ", ".join([str(x.IntegerValue) for x in self.host_blocker_ids]))
        if self.link_blockers:
            lnk_txt = []
            for (lid, leid) in self.link_blockers:
                lnk_txt.append("L{0}:E{1}".format(lid.IntegerValue, leid.IntegerValue))
            parts.append("Link: " + ", ".join(lnk_txt))
        return " | ".join(parts) if parts else "-"

# ----------------------------
# Equipment classification
# ----------------------------
def is_transformer(fi):
    """Проверка, является ли элемент трансформатором"""
    try:
        # 1. Проверяем параметр SHN_EquipmentClass
        cls = get_param_str(fi, SHN_CLASS_PARAM)
        if cls:
            cls_up = cls.upper()
            if "TR" in cls_up:
                return True
        
        # 2. Проверяем имя типа/семейства
        try:
            fam_name = fi.Symbol.FamilyName.lower()
            type_name = fi.Symbol.get_Parameter(BuiltInCategory.OST_ElectricalEquipment).AsValueString().lower()
            combined = fam_name + " " + type_name
        except:
            try:
                fam_name = fi.Symbol.FamilyName.lower()
                combined = fam_name
            except:
                combined = ""
        
        for key in TRANSFORMER_NAME_KEYS:
            if key.lower() in combined:
                return True
    except:
        pass
    return False

def is_electrical_cabinet(fi):
    """Проверка, является ли элемент шкафом/щитом"""
    try:
        # Если это трансформатор - не шкаф
        if is_transformer(fi):
            return False
        
        # Проверяем параметр SHN_EquipmentClass
        cls = get_param_str(fi, SHN_CLASS_PARAM)
        if cls:
            cls_up = cls.upper()
            if "LV" in cls_up or "MV" in cls_up:
                return True
        
        # По умолчанию считаем все не-трансформаторы электрооборудования шкафами
        return True
    except:
        pass
    return False

# ----------------------------
# Collection functions
# ----------------------------
def collect_family_instances_by_bic(bic):
    """Сбор FamilyInstance по BuiltInCategory"""
    try:
        collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        return list(collector)
    except:
        return []

def collect_host_obstacles():
    """Сбор препятствий в хост-модели"""
    try:
        # Используем ElementMulticategoryFilter вместо LogicalOrFilter
        bic_list = List[BuiltInCategory](HOST_OBSTACLE_BIC)
        multi_filter = ElementMulticategoryFilter(bic_list)
        collector = FilteredElementCollector(doc).WherePasses(multi_filter).WhereElementIsNotElementType()
        return list(collector)
    except Exception as ex:
        logger.warning("collect_host_obstacles error: {0}".format(ex))
        return []

def collect_links():
    """Сбор всех RevitLinkInstance"""
    try:
        collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
        links = []
        for link in collector:
            try:
                link_doc = link.GetLinkDocument()
                if link_doc:
                    links.append(link)
            except:
                continue
        return links
    except:
        return []

def collect_obstacles_from_link(link_doc):
    """Сбор препятствий из связанного документа"""
    try:
        # Используем ElementMulticategoryFilter
        bic_list = List[BuiltInCategory](LINK_OBSTACLE_BIC)
        multi_filter = ElementMulticategoryFilter(bic_list)
        collector = FilteredElementCollector(link_doc).WherePasses(multi_filter).WhereElementIsNotElementType()
        return list(collector)
    except Exception as ex:
        logger.warning("collect_obstacles_from_link error: {0}".format(ex))
        return []

# ----------------------------
# Service zone creation
# ----------------------------
def create_service_zone_solid(cabinet_fi, depth_mm):
    """
    Создание зоны обслуживания перед шкафом
    depth_mm - глубина зоны в мм (например 800)
    Возвращает solid в координатах модели
    """
    try:
        # Получаем facing orientation
        facing = cabinet_fi.FacingOrientation
        if not facing:
            return None
        
        facing = normalize(facing)
        if not facing:
            return None
        
        # Получаем bbox шкафа
        bb = try_get_bbox(cabinet_fi, None)
        if not bb:
            return None
        
        # Локальная система координат
        location = cabinet_fi.Location
        if not hasattr(location, 'Point'):
            return None
        origin = location.Point
        
        # Получаем размеры шкафа (упрощенно)
        width = bb.Max.X - bb.Min.X
        height = bb.Max.Z - bb.Min.Z
        
        # Создаем локальную ось Y перпендикулярную facing
        if abs(facing.Z) < 0.9:
            up = XYZ.BasisZ
        else:
            up = XYZ.BasisX
        
        right = facing.CrossProduct(up)
        right = normalize(right)
        if not right:
            return None
        
        up = right.CrossProduct(facing)
        up = normalize(up)
        
        # Создаем transform для локальной системы
        trf = Transform.Identity
        trf.Origin = origin
        trf.BasisX = right
        trf.BasisY = facing
        trf.BasisZ = up
        
        # Создаем зону в локальных координатах
        depth_ft = mm_to_ft(depth_mm)
        half_width = width / 2.0
        
        # Зона: от -half_width до +half_width по X, от 0 до depth_ft по Y, от 0 до height по Z
        local_solid = make_local_box_solid(-half_width, half_width, 0, depth_ft, 0, height)
        if not local_solid:
            return None
        
        # Трансформируем в мировые координаты
        global_solid = SolidUtils.CreateTransformed(local_solid, trf)
        return global_solid
    except Exception as ex:
        logger.warning("create_service_zone_solid error: {0}".format(ex))
        return None

def create_clearance_zone_bbox(equipment_fi, clearance_mm):
    """
    Создание зоны зазора вокруг оборудования (для трансформаторов)
    Возвращает BoundingBoxXYZ расширенный на clearance_mm
    """
    try:
        bb = try_get_bbox(equipment_fi, None)
        if not bb:
            return None
        
        delta_ft = mm_to_ft(clearance_mm)
        expanded = bbox_expand(bb, delta_ft)
        return expanded
    except:
        return None

# ----------------------------
# CHECK 1: Cabinet clash with linked models
# ----------------------------
def check_1_cabinet_clash_links(cabinets, links, geom_opt):
    """
    Проверка столкновения шкафов с элементами связанных моделей
    """
    issues = []
    
    for cab in cabinets:
        try:
            cab_bb = try_get_bbox(cab, None)
            if not cab_bb:
                continue
            
            cab_solids = get_solids(cab, geom_opt)
            if not cab_solids:
                continue
            
            # Проверяем каждую связь
            for link in links:
                try:
                    link_doc = link.GetLinkDocument()
                    if not link_doc:
                        continue
                    
                    link_transform = link.GetTotalTransform()
                    obstacles = collect_obstacles_from_link(link_doc)
                    
                    for obs in obstacles:
                        try:
                            obs_bb = try_get_bbox(obs, None)
                            if not obs_bb:
                                continue
                            
                            # Трансформируем bbox препятствия в координаты хоста
                            obs_bb_host = transform_bbox(obs_bb, link_transform)
                            
                            # Быстрая проверка bbox
                            if not bbox_intersects(cab_bb, obs_bb_host):
                                continue
                            
                            # Детальная проверка solid
                            obs_solids = get_solids(obs, geom_opt, link_transform)
                            if solids_intersect(cab_solids, obs_solids):
                                # Создаем issue
                                msg = "Cabinet clashes with linked element (Link: {0}, Element: {1})".format(
                                    link.Name, obs.Id.IntegerValue
                                )
                                issue = Issue(
                                    Severity.CRITICAL,
                                    "1) Clash vs linked models",
                                    cab.Id,
                                    msg,
                                    bbox_union(cab_bb, obs_bb_host)
                                )
                                issue.add_link_blocker(link.Id, obs.Id)
                                issues.append(issue)
                        except Exception as ex:
                            logger.debug("check_1 obstacle error: {0}".format(ex))
                            continue
                except Exception as ex:
                    logger.warning("check_1 link error: {0}".format(ex))
                    continue
        except Exception as ex:
            logger.warning("check_1 cabinet error: {0}".format(ex))
            continue
    
    return issues

# ----------------------------
# CHECK 2: LV service zone 800mm
# ----------------------------
def check_2_lv_service_zone(cabinets, host_obstacles, links, geom_opt):
    """
    Проверка зоны обслуживания 800мм перед LV шкафами
    """
    issues = []
    
    for cab in cabinets:
        try:
            # Создаем зону обслуживания
            service_solid = create_service_zone_solid(cab, SERVICE_MM)
            if not service_solid:
                continue
            
            service_bb = service_solid.GetBoundingBox()
            blockers_found = []
            
            # Проверяем препятствия в хосте
            for obs in host_obstacles:
                try:
                    # Не проверяем сам шкаф
                    if obs.Id == cab.Id:
                        continue
                    
                    obs_bb = try_get_bbox(obs, None)
                    if not obs_bb:
                        continue
                    
                    if not bbox_intersects(service_bb, obs_bb):
                        continue
                    
                    obs_solids = get_solids(obs, geom_opt)
                    if solids_intersect([service_solid], obs_solids):
                        blockers_found.append(("host", obs.Id, None))
                except:
                    continue
            
            # Проверяем препятствия в связях
            for link in links:
                try:
                    link_doc = link.GetLinkDocument()
                    if not link_doc:
                        continue
                    
                    link_transform = link.GetTotalTransform()
                    link_obstacles = collect_obstacles_from_link(link_doc)
                    
                    for obs in link_obstacles:
                        try:
                            obs_bb = try_get_bbox(obs, None)
                            if not obs_bb:
                                continue
                            
                            obs_bb_host = transform_bbox(obs_bb, link_transform)
                            if not bbox_intersects(service_bb, obs_bb_host):
                                continue
                            
                            obs_solids = get_solids(obs, geom_opt, link_transform)
                            if solids_intersect([service_solid], obs_solids):
                                blockers_found.append(("link", link.Id, obs.Id))
                        except:
                            continue
                except:
                    continue
            
            # Если найдены блокирующие элементы - создаем issue
            if blockers_found:
                msg = "Service zone (800mm) blocked by {0} element(s)".format(len(blockers_found))
                issue = Issue(
                    Severity.CRITICAL,
                    "2) LV service zone 800mm",
                    cab.Id,
                    msg,
                    bbox_expand(service_bb, mm_to_ft(100))  # немного расширяем для визуализации
                )
                
                for (btype, id1, id2) in blockers_found:
                    if btype == "host":
                        issue.add_host_blocker(id1)
                    else:
                        issue.add_link_blocker(id1, id2)
                
                issues.append(issue)
                
        except Exception as ex:
            logger.warning("check_2 cabinet error: {0}".format(ex))
            continue
    
    return issues

# ----------------------------
# ----------------------------
# CHECK 3: Double-row corridor clearance
# ----------------------------
def get_element_room(element):
    """
    Получить Room в котором находится элемент
    Пробует несколько способов
    """
    try:
        # Способ 1: Через свойство Room (для FamilyInstance)
        if hasattr(element, 'Room') and element.Room:
            return element.Room.Id.IntegerValue
    except:
        pass
    
    try:
        # Способ 2: Через параметр Room Name
        room_param = element.get_Parameter(BuiltInCategory.ROOM_NAME)
        if room_param and room_param.AsString():
            return room_param.AsString()
    except:
        pass
    
    try:
        # Способ 3: Spatial calculation - найти Room по точке
        location = element.Location
        if hasattr(location, 'Point'):
            point = location.Point
            
            # Ищем Room содержащий эту точку
            rooms_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
            for room in rooms_collector:
                try:
                    if room.IsPointInRoom(point):
                        return room.Id.IntegerValue
                except:
                    continue
    except:
        pass
    
    return None

def check_wall_between_cabinets(pos_a, pos_b, links):
    """
    Проверка наличия стены между двумя шкафами в связанных моделях
    Возвращает True если есть стена (НЕ коридор)
    Возвращает False если нет стены (коридор)
    """
    try:
        # Создаем линию между центрами шкафов
        line_vec = pos_b - pos_a
        line_length = line_vec.GetLength()
        
        # Проверяем каждую связь
        for link in links:
            try:
                link_doc = link.GetLinkDocument()
                if not link_doc:
                    continue
                
                link_transform = link.GetTotalTransform()
                
                # Трансформируем позиции в координаты связи
                try:
                    inv_transform = link_transform.Inverse
                    pos_a_link = inv_transform.OfPoint(pos_a)
                    pos_b_link = inv_transform.OfPoint(pos_b)
                except:
                    # Если не удалось инвертировать - пропускаем
                    continue
                
                # Собираем стены из связи
                walls_collector = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
                
                for wall in walls_collector:
                    try:
                        # Получаем curve стены
                        loc_curve = wall.Location
                        if not hasattr(loc_curve, 'Curve'):
                            continue
                        
                        wall_curve = loc_curve.Curve
                        
                        # Проверяем пересечение методом проекции
                        # Находим ближайшую точку на линии стены к середине линии между шкафами
                        mid_point = XYZ(
                            (pos_a_link.X + pos_b_link.X) / 2,
                            (pos_a_link.Y + pos_b_link.Y) / 2,
                            (pos_a_link.Z + pos_b_link.Z) / 2
                        )
                        
                        # Проецируем середину на curve стены
                        result = wall_curve.Project(mid_point)
                        
                        if result:
                            # Проверяем что точка проекции лежит на отрезке между шкафами
                            proj_point = result.XYZPoint
                            
                            # Расстояние от проекции до обоих шкафов
                            dist_to_a = proj_point.DistanceTo(pos_a_link)
                            dist_to_b = proj_point.DistanceTo(pos_b_link)
                            
                            # Если сумма расстояний примерно равна длине линии - стена между шкафами
                            total_dist = dist_to_a + dist_to_b
                            
                            # Допуск 10% на неточности
                            if total_dist < line_length * 1.1:
                                # Стена пересекает линию между шкафами
                                logger.info("    → WALL FOUND between cabinets: Wall ID={0} in link {1} (distance check passed)".format(
                                    wall.Id.IntegerValue,
                                    link.Id.IntegerValue
                                ))
                                return True
                            
                    except Exception as ex:
                        logger.debug("    Wall check error for wall {0}: {1}".format(
                            wall.Id.IntegerValue if hasattr(wall, 'Id') else '?',
                            ex
                        ))
                        continue
                        
            except Exception as ex:
                logger.debug("check_wall_between error for link: {0}".format(ex))
                continue
        
        logger.debug("    → NO WALL between cabinets")
        return False
        
    except Exception as ex:
        logger.debug("check_wall_between error: {0}".format(ex))
        return False

def check_cabinets_in_same_room(cab_a, cab_b):
    """
    Проверка что два шкафа в одном помещении
    Возвращает True если в одном помещении (или не удалось определить)
    Возвращает False если точно в разных помещениях
    """
    try:
        room_a = get_element_room(cab_a)
        room_b = get_element_room(cab_b)
        
        # Если не удалось определить хотя бы одно помещение - считаем что в одном
        # (чтобы не пропустить потенциальную ошибку)
        if room_a is None or room_b is None:
            logger.debug("Cannot determine room for cabinets {0} or {1} - assuming same room".format(
                cab_a.Id.IntegerValue, cab_b.Id.IntegerValue
            ))
            return True
        
        # Сравниваем помещения
        same_room = (room_a == room_b)
        
        if not same_room:
            logger.info("Cabinets {0} and {1} are in DIFFERENT rooms ({2} vs {3}) - skipping corridor check".format(
                cab_a.Id.IntegerValue,
                cab_b.Id.IntegerValue,
                room_a,
                room_b
            ))
        else:
            logger.debug("Cabinets {0} and {1} are in SAME room ({2})".format(
                cab_a.Id.IntegerValue,
                cab_b.Id.IntegerValue,
                room_a
            ))
        
        return same_room
        
    except Exception as ex:
        logger.debug("check_cabinets_in_same_room error: {0} - assuming same room".format(ex))
        # При ошибке считаем что в одном помещении (не пропускаем проверку)
        return True

def check_3_double_row_corridor(cabinets, host_obstacles=None, links=None, geom_opt=None):
    """
    Проверка зазора между рядами шкафов "лицом к лицу"
    Требуемое расстояние = Depth1 + Depth2 + 800mm
    УЛУЧШЕНО: Проверяет что шкафы в одном помещении (через Room параметр)
    """
    issues = []
    
    search_radius_ft = mm_to_ft(DOUBLE_ROW_SEARCH_RADIUS_MM)
    corridor_ft = mm_to_ft(CORRIDOR_CLEARANCE_MM)
    
    checked_pairs = 0
    skipped_different_rooms = 0
    found_issues = 0
    
    for i, cab_a in enumerate(cabinets):
        try:
            facing_a = cab_a.FacingOrientation
            if not facing_a:
                continue
            facing_a = normalize(facing_a)
            if not facing_a:
                continue
            
            loc_a = cab_a.Location
            if not hasattr(loc_a, 'Point'):
                continue
            pos_a = loc_a.Point
            
            bb_a = try_get_bbox(cab_a, None)
            if not bb_a:
                continue
            
            # Примерная глубина шкафа A (вдоль facing)
            size_a = bb_a.Max - bb_a.Min
            depth_a = abs(facing_a.DotProduct(XYZ(size_a.X, size_a.Y, size_a.Z)))
            
            # Ищем кандидатов поблизости
            for j, cab_b in enumerate(cabinets):
                if j <= i:  # избегаем дублей
                    continue
                
                try:
                    loc_b = cab_b.Location
                    if not hasattr(loc_b, 'Point'):
                        continue
                    pos_b = loc_b.Point
                    
                    dist = pos_a.DistanceTo(pos_b)
                    if dist > search_radius_ft:
                        continue
                    
                    facing_b = cab_b.FacingOrientation
                    if not facing_b:
                        continue
                    facing_b = normalize(facing_b)
                    if not facing_b:
                        continue
                    
                    # Проверяем что они смотрят друг на друга (dot < -0.85 - более строго!)
                    dot = facing_a.DotProduct(facing_b)
                    if dot > -0.85:  # Было -0.75, стало -0.85 (только почти точно напротив)
                        continue
                    
                    # НОВАЯ ПРОВЕРКА: Шкафы должны быть выровнены (не по диагонали)
                    # Вектор между шкафами должен быть примерно параллелен facing_a
                    vec_ab = pos_b - pos_a
                    vec_ab_norm = normalize(vec_ab)
                    if not vec_ab_norm:
                        continue
                    
                    # Проверяем что вектор AB параллелен facing_a (или -facing_a)
                    alignment = abs(vec_ab_norm.DotProduct(facing_a))
                    
                    # alignment должен быть близок к 1.0 (параллельны)
                    # Если < 0.85, значит шкафы стоят под углом (по диагонали)
                    if alignment < 0.85:
                        logger.debug("→ SKIPPED (not aligned): alignment={0:.3f} (diagonal placement)".format(alignment))
                        continue
                    
                    checked_pairs += 1
                    
                    logger.info("Check 3: Found facing pair #{0} - Cabinet A: {1}, Cabinet B: {2}, dot={3:.3f}, alignment={4:.3f}, distance={5:.2f}ft ({6:.0f}mm)".format(
                        checked_pairs,
                        cab_a.Id.IntegerValue,
                        cab_b.Id.IntegerValue,
                        dot,
                        alignment,
                        dist,
                        dist * MM_IN_FT
                    ))
                    
                    # ПРОВЕРКА: Шкафы в одном помещении?
                    if not check_cabinets_in_same_room(cab_a, cab_b):
                        # В разных помещениях - пропускаем
                        skipped_different_rooms += 1
                        logger.info("→ SKIPPED (different rooms)")
                        continue
                    
                    logger.info("→ Same room - checking for wall between cabinets")
                    
                    # НОВАЯ ПРОВЕРКА: Есть ли стена между шкафами?
                    if check_wall_between_cabinets(pos_a, pos_b, links):
                        # Есть стена между шкафами - это не коридор!
                        logger.info("→ SKIPPED (wall between cabinets - not a corridor)")
                        continue
                    
                    logger.info("→ No wall between - checking corridor clearance")
                    
                    bb_b = try_get_bbox(cab_b, None)
                    if not bb_b:
                        continue
                    
                    size_b = bb_b.Max - bb_b.Min
                    depth_b = abs(facing_b.DotProduct(XYZ(size_b.X, size_b.Y, size_b.Z)))
                    
                    # Расстояние вдоль facing_a между центрами (vec_ab уже вычислен выше)
                    dist_along = abs(vec_ab.DotProduct(facing_a))
                    
                    # Требуемое расстояние
                    required = depth_a + depth_b + corridor_ft
                    
                    if dist_along < required:
                        # Нарушение!
                        found_issues += 1
                        shortage_mm = (required - dist_along) * MM_IN_FT
                        msg = "Double-row corridor clearance insufficient (same room). Required: {0:.0f}mm, Actual: {1:.0f}mm, Shortage: {2:.0f}mm".format(
                            required * MM_IN_FT,
                            dist_along * MM_IN_FT,
                            shortage_mm
                        )
                        
                        logger.info("→ ISSUE FOUND: {0}".format(msg))
                        
                        issue = Issue(
                            Severity.CRITICAL,
                            "3) Double-row corridor",
                            cab_a.Id,
                            msg,
                            bbox_union(bb_a, bb_b)
                        )
                        issue.add_host_blocker(cab_b.Id)
                        issues.append(issue)
                    else:
                        logger.info("→ OK: Clearance sufficient ({0:.0f}mm)".format(dist_along * MM_IN_FT))
                        
                except Exception as ex:
                    logger.debug("check_3 inner loop error: {0}".format(ex))
                    continue
        except Exception as ex:
            logger.warning("check_3 cabinet error: {0}".format(ex))
            continue
    
    logger.info("Check 3 summary: {0} facing pairs found, {1} skipped (different rooms), {2} issues found".format(
        checked_pairs,
        skipped_different_rooms,
        found_issues
    ))
    
    return issues

# ----------------------------
# CHECK 4: Transformer clearance zones
# ----------------------------

# ----------------------------
# CHECK 4: Transformer clearance zones
# ----------------------------
def check_4_transformer_clearance(transformers, host_obstacles, links, geom_opt):
    """
    Проверка зазоров вокруг трансформаторов:
    Проверяется что вокруг трансформатора (в радиусе 800мм/1000мм и высоте до +1000мм)
    НЕТ стен и конструкций из СВЯЗАННЫХ моделей (архитектура/конструкции)
    
    - < 800mm = CRITICAL
    - 800-1000mm = WARNING  
    - > 1000mm = OK
    
    ВАЖНО: Проверяются ТОЛЬКО связанные модели, НЕ хост-модель!
    """
    issues = []
    
    # Используем глобальную конфигурацию категорий
    link_blocking_categories = CHECK4_LINK_CATEGORIES
    
    for tr in transformers:
        try:
            logger.info("=" * 60)
            logger.info("Check 4: Checking transformer ID={0}".format(tr.Id.IntegerValue))
            
            tr_bb = try_get_bbox(tr, None)
            if not tr_bb:
                logger.warning("Cannot get bbox for transformer {0}".format(tr.Id.IntegerValue))
                continue
            
            # Получаем центр трансформатора
            tr_center = XYZ(
                (tr_bb.Min.X + tr_bb.Max.X) / 2,
                (tr_bb.Min.Y + tr_bb.Max.Y) / 2,
                (tr_bb.Min.Z + tr_bb.Max.Z) / 2
            )
            
            # Размеры трансформатора (в плане - X, Y)
            tr_width_x = tr_bb.Max.X - tr_bb.Min.X
            tr_width_y = tr_bb.Max.Y - tr_bb.Min.Y
            tr_height = tr_bb.Max.Z - tr_bb.Min.Z
            
            logger.info("Transformer center: ({0:.2f}, {1:.2f}, {2:.2f})".format(
                tr_center.X, tr_center.Y, tr_center.Z
            ))
            logger.info("Transformer size: {0:.0f}mm x {1:.0f}mm x {2:.0f}mm (height)".format(
                tr_width_x * MM_IN_FT,
                tr_width_y * MM_IN_FT,
                tr_height * MM_IN_FT
            ))
            
            # Создаем зоны проверки вокруг трансформатора
            # CRITICAL зона: +800мм вокруг, высота от пола до +TRANS_HEIGHT_CHECK_MM над трансформатором
            crit_radius_ft = mm_to_ft(TRANS_CRIT_MM)
            warn_radius_ft = mm_to_ft(TRANS_WARN_MM)
            height_above_ft = mm_to_ft(TRANS_HEIGHT_CHECK_MM)  # Настраиваемая высота проверки
            
            # Зона CRITICAL
            zone_crit_min = XYZ(
                tr_bb.Min.X - crit_radius_ft,
                tr_bb.Min.Y - crit_radius_ft,
                tr_bb.Min.Z  # от уровня пола трансформатора
            )
            zone_crit_max = XYZ(
                tr_bb.Max.X + crit_radius_ft,
                tr_bb.Max.Y + crit_radius_ft,
                tr_bb.Max.Z + height_above_ft  # +1м над трансформатором
            )
            
            # Зона WARNING
            zone_warn_min = XYZ(
                tr_bb.Min.X - warn_radius_ft,
                tr_bb.Min.Y - warn_radius_ft,
                tr_bb.Min.Z
            )
            zone_warn_max = XYZ(
                tr_bb.Max.X + warn_radius_ft,
                tr_bb.Max.Y + warn_radius_ft,
                tr_bb.Max.Z + height_above_ft
            )
            
            zone_crit_bb = BoundingBoxXYZ()
            zone_crit_bb.Min = zone_crit_min
            zone_crit_bb.Max = zone_crit_max
            
            zone_warn_bb = BoundingBoxXYZ()
            zone_warn_bb.Min = zone_warn_min
            zone_warn_bb.Max = zone_warn_max
            
            logger.info("Critical zone: radius={0}mm, height={1}mm above transformer".format(
                TRANS_CRIT_MM,
                TRANS_HEIGHT_CHECK_MM
            ))
            logger.info("Warning zone: radius={0}mm, height={1}mm above transformer".format(
                TRANS_WARN_MM,
                TRANS_HEIGHT_CHECK_MM
            ))
            
            crit_blockers = []
            warn_blockers = []
            
            total_obstacles_checked = 0
            
            # ========================================
            # ПРОВЕРЯЕМ ТОЛЬКО СВЯЗАННЫЕ МОДЕЛИ!
            # ========================================
            logger.info("Checking linked models...")
            
            for link in links:
                try:
                    link_doc = link.GetLinkDocument()
                    if not link_doc:
                        continue
                    
                    link_name = link.Name
                    logger.info("  Checking link: {0}".format(link_name))
                    
                    link_transform = link.GetTotalTransform()
                    
                    # Собираем стены и конструкции из связи
                    bic_list = List[BuiltInCategory](link_blocking_categories)
                    multi_filter = ElementMulticategoryFilter(bic_list)
                    link_obs_collector = FilteredElementCollector(link_doc).WherePasses(multi_filter).WhereElementIsNotElementType()
                    
                    link_obstacles = list(link_obs_collector)
                    logger.info("    Found {0} potential obstacles in link".format(len(link_obstacles)))
                    
                    for obs in link_obstacles:
                        try:
                            total_obstacles_checked += 1
                            
                            obs_bb = try_get_bbox(obs, None)
                            if not obs_bb:
                                continue
                            
                            # Трансформируем bbox в координаты хоста
                            obs_bb_host = transform_bbox(obs_bb, link_transform)
                            
                            # Быстрая проверка - далеко ли вообще элемент
                            # Используем расширенную warning зону для первичной фильтрации
                            quick_check_zone = BoundingBoxXYZ()
                            margin = mm_to_ft(500)  # +500мм для фильтрации
                            quick_check_zone.Min = XYZ(
                                zone_warn_bb.Min.X - margin,
                                zone_warn_bb.Min.Y - margin,
                                zone_warn_bb.Min.Z - margin
                            )
                            quick_check_zone.Max = XYZ(
                                zone_warn_bb.Max.X + margin,
                                zone_warn_bb.Max.Y + margin,
                                zone_warn_bb.Max.Z + margin
                            )
                            
                            if not bbox_intersects(quick_check_zone, obs_bb_host):
                                continue
                            
                            # Получаем категорию для логирования
                            obs_cat_name = "Unknown"
                            try:
                                if obs.Category:
                                    obs_cat_name = obs.Category.Name
                            except:
                                pass
                            
                            # ============================================
                            # ВАЖНО: Используем SOLID вместо BBOX!
                            # ============================================
                            
                            # Получаем solid препятствия
                            obs_geom_opt = Options()
                            obs_geom_opt.DetailLevel = ViewDetailLevel.Fine
                            obs_geom_opt.IncludeNonVisibleObjects = False  # ТОЛЬКО видимая геометрия!
                            
                            obs_solids = []
                            try:
                                obs_geom_elem = obs.get_Geometry(obs_geom_opt)
                                if obs_geom_elem:
                                    for geom_obj in obs_geom_elem:
                                        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0.0001:
                                            obs_solids.append(geom_obj)
                                        elif isinstance(geom_obj, GeometryInstance):
                                            inst_geom = geom_obj.GetInstanceGeometry()
                                            if inst_geom:
                                                for inst_obj in inst_geom:
                                                    if isinstance(inst_obj, Solid) and inst_obj.Volume > 0.0001:
                                                        obs_solids.append(inst_obj)
                            except:
                                pass
                            
                            if not obs_solids:
                                # Если нет solid - используем bbox как fallback
                                logger.debug("    Obstacle {0} (ID={1}): No solids, using bbox".format(
                                    obs_cat_name, obs.Id.IntegerValue
                                ))
                                obs_center_local = XYZ(
                                    (obs_bb.Min.X + obs_bb.Max.X) / 2,
                                    (obs_bb.Min.Y + obs_bb.Max.Y) / 2,
                                    (obs_bb.Min.Z + obs_bb.Max.Z) / 2
                                )
                                obs_center_host = link_transform.OfPoint(obs_center_local)
                                
                                dist_plan = math.sqrt(
                                    (obs_center_host.X - tr_center.X) ** 2 +
                                    (obs_center_host.Y - tr_center.Y) ** 2
                                )
                                dist_plan_mm = dist_plan * MM_IN_FT
                                
                                obs_size = obs_bb.Max - obs_bb.Min
                                obs_radius_ft = max(obs_size.X, obs_size.Y) / 2
                                tr_radius_ft = max(tr_width_x, tr_width_y) / 2
                                
                                clearance_ft = dist_plan - tr_radius_ft - obs_radius_ft
                                clearance_mm = clearance_ft * MM_IN_FT
                            else:
                                # Используем SOLID для точного расстояния!
                                # Трансформируем solid в координаты хоста
                                obs_solids_host = []
                                for solid in obs_solids:
                                    try:
                                        solid_host = SolidUtils.CreateTransformed(solid, link_transform)
                                        obs_solids_host.append(solid_host)
                                    except:
                                        continue
                                
                                if not obs_solids_host:
                                    continue
                                
                                # Находим ближайшую точку на solid препятствия к центру трансформатора
                                min_distance = float('inf')
                                
                                for solid_host in obs_solids_host:
                                    try:
                                        # Получаем все грани solid
                                        for face in solid_host.Faces:
                                            try:
                                                # Проецируем центр трансформатора на грань
                                                result = face.Project(tr_center)
                                                if result:
                                                    point_on_face = result.XYZPoint
                                                    # Расстояние в плане (2D)
                                                    dist_plan = math.sqrt(
                                                        (point_on_face.X - tr_center.X) ** 2 +
                                                        (point_on_face.Y - tr_center.Y) ** 2
                                                    )
                                                    if dist_plan < min_distance:
                                                        min_distance = dist_plan
                                            except:
                                                continue
                                    except:
                                        continue
                                
                                if min_distance == float('inf'):
                                    # Fallback - используем bbox
                                    obs_center_local = XYZ(
                                        (obs_bb.Min.X + obs_bb.Max.X) / 2,
                                        (obs_bb.Min.Y + obs_bb.Max.Y) / 2,
                                        (obs_bb.Min.Z + obs_bb.Max.Z) / 2
                                    )
                                    obs_center_host = link_transform.OfPoint(obs_center_local)
                                    dist_plan = math.sqrt(
                                        (obs_center_host.X - tr_center.X) ** 2 +
                                        (obs_center_host.Y - tr_center.Y) ** 2
                                    )
                                    obs_size = obs_bb.Max - obs_bb.Min
                                    obs_radius_ft = max(obs_size.X, obs_size.Y) / 2
                                    tr_radius_ft = max(tr_width_x, tr_width_y) / 2
                                    clearance_ft = dist_plan - tr_radius_ft - obs_radius_ft
                                else:
                                    # min_distance - это расстояние от центра TR до ближайшей точки на solid
                                    # Вычитаем только радиус трансформатора
                                    tr_radius_ft = max(tr_width_x, tr_width_y) / 2
                                    clearance_ft = min_distance - tr_radius_ft
                                
                                dist_plan_mm = min_distance * MM_IN_FT if min_distance != float('inf') else 0
                                clearance_mm = clearance_ft * MM_IN_FT
                            
                            # ДЕТАЛЬНОЕ логирование для отладки
                            logger.info("    >>> Obstacle {0} (ID={1}):".format(obs_cat_name, obs.Id.IntegerValue))
                            logger.info("        Method: {0}".format("SOLID (visible geometry)" if obs_solids else "BBOX (fallback)"))
                            logger.info("        CLEARANCE (surface-to-surface): {0:.0f}mm".format(clearance_mm))
                            
                            # Проверка CRITICAL зоны: < 800mm (НЕ <=)
                            if clearance_mm < TRANS_CRIT_MM:
                                logger.info("        → CRITICAL: {0:.0f}mm < {1}mm".format(clearance_mm, TRANS_CRIT_MM))
                                crit_blockers.append(("link", link.Id, obs.Id))
                                continue
                            
                            # Проверка WARNING зоны: >= 800mm AND < 1000mm
                            if clearance_mm >= TRANS_CRIT_MM and clearance_mm < TRANS_WARN_MM:
                                logger.info("        → WARNING: {0:.0f}mm in range {1}-{2}mm".format(
                                    clearance_mm,
                                    TRANS_CRIT_MM,
                                    TRANS_WARN_MM
                                ))
                                warn_blockers.append(("link", link.Id, obs.Id))
                                continue
                            
                            # >= 1000mm - OK
                            logger.info("        → OK: {0:.0f}mm >= {1}mm".format(clearance_mm, TRANS_WARN_MM))
                                
                        except Exception as ex:
                            logger.debug("    Error checking obstacle: {0}".format(ex))
                            continue
                    
                except Exception as ex:
                    logger.warning("  Error checking link {0}: {1}".format(link_name if 'link_name' in locals() else '?', ex))
                    continue
            
            logger.info("Total obstacles checked: {0}".format(total_obstacles_checked))
            logger.info("Critical zone violations: {0}".format(len(crit_blockers)))
            logger.info("Warning zone violations: {0}".format(len(warn_blockers)))
            
            # Создаем issues
            if crit_blockers:
                # Собираем информацию о блокирующих элементах для сообщения
                blocker_details = []
                for (btype, link_id, obs_id) in crit_blockers:
                    try:
                        link = doc.GetElement(link_id)
                        link_doc = link.GetLinkDocument() if link else None
                        if link_doc:
                            obs_elem = link_doc.GetElement(obs_id)
                            if obs_elem and obs_elem.Category:
                                cat_name = obs_elem.Category.Name
                                blocker_details.append("{0} (ID:{1})".format(cat_name, obs_id.IntegerValue))
                            else:
                                blocker_details.append("Element ID:{0}".format(obs_id.IntegerValue))
                        else:
                            blocker_details.append("Element ID:{0}".format(obs_id.IntegerValue))
                    except:
                        blocker_details.append("Element ID:{0}".format(obs_id.IntegerValue))
                
                msg = "Transformer clearance CRITICAL (<800mm): {0} linked obstacle(s) found: {1}".format(
                    len(crit_blockers),
                    ", ".join(blocker_details) if blocker_details else "see logs"
                )
                issue = Issue(
                    Severity.CRITICAL,
                    "4) Transformer clearance 800/1000",
                    tr.Id,
                    msg,
                    zone_crit_bb
                )
                for (btype, id1, id2) in crit_blockers:
                    issue.add_link_blocker(id1, id2)
                issues.append(issue)
                logger.info("→ ISSUE CREATED (CRITICAL): {0}".format(msg))
                
            elif warn_blockers:
                # Собираем информацию о блокирующих элементах для сообщения
                blocker_details = []
                for (btype, link_id, obs_id) in warn_blockers:
                    try:
                        link = doc.GetElement(link_id)
                        link_doc = link.GetLinkDocument() if link else None
                        if link_doc:
                            obs_elem = link_doc.GetElement(obs_id)
                            if obs_elem and obs_elem.Category:
                                cat_name = obs_elem.Category.Name
                                blocker_details.append("{0} (ID:{1})".format(cat_name, obs_id.IntegerValue))
                            else:
                                blocker_details.append("Element ID:{0}".format(obs_id.IntegerValue))
                        else:
                            blocker_details.append("Element ID:{0}".format(obs_id.IntegerValue))
                    except:
                        blocker_details.append("Element ID:{0}".format(obs_id.IntegerValue))
                
                msg = "Transformer clearance WARNING (800-1000mm): {0} linked obstacle(s) found: {1}".format(
                    len(warn_blockers),
                    ", ".join(blocker_details) if blocker_details else "see logs"
                )
                issue = Issue(
                    Severity.WARNING,
                    "4) Transformer clearance 800/1000",
                    tr.Id,
                    msg,
                    zone_warn_bb
                )
                for (btype, id1, id2) in warn_blockers:
                    issue.add_link_blocker(id1, id2)
                issues.append(issue)
                logger.info("→ ISSUE CREATED (WARNING): {0}".format(msg))
            else:
                logger.info("→ OK: No clearance issues (no linked obstacles in zones)")
                
        except Exception as ex:
            logger.error("check_4 transformer error: {0}".format(ex))
            import traceback
            logger.error(traceback.format_exc())
            continue
    
    logger.info("=" * 60)
    logger.info("Check 4 completed: {0} transformer(s) checked, {1} issue(s) found".format(
        len(transformers),
        len(issues)
    ))
    
    return issues

class NavHandler(IExternalEventHandler):
    def __init__(self):
        self.command = None
    
    def Execute(self, uiapp):
        try:
            if not self.command:
                return
            
            cmd_type = self.command.get("type")
            issue = self.command.get("issue")
            
            if not issue:
                return
            
            if cmd_type == "select":
                self.do_select(issue)
            elif cmd_type == "3d":
                self.do_3d_section(issue)
                
        except Exception as ex:
            logger.error("NavHandler.Execute error: {0}".format(ex))
        finally:
            self.command = None
    
    def GetName(self):
        return "SHN_NavHandler"
    
    def do_select(self, issue):
        """Выделить элементы issue"""
        try:
            ids_to_select = List[ElementId]()
            
            # Главный элемент
            ids_to_select.Add(issue.main_id)
            
            # Host blockers
            for hid in issue.host_blocker_ids:
                ids_to_select.Add(hid)
            
            # Link instances (нельзя выделить элементы внутри связей напрямую, но можно выделить саму связь)
            for (link_id, _) in issue.link_blockers:
                if link_id not in [x for x in ids_to_select]:
                    ids_to_select.Add(link_id)
            
            uidoc.Selection.SetElementIds(ids_to_select)
            
            # Зум на главный элемент
            try:
                uidoc.ShowElements(issue.main_id)
            except:
                pass
                
        except Exception as ex:
            logger.error("do_select error: {0}".format(ex))
    
    def do_3d_section(self, issue):
        """Открыть 3D вид с Section Box вокруг issue"""
        try:
            # Найти или создать 3D вид
            view3d = self.find_or_create_3d_view()
            if not view3d:
                forms.alert("Cannot create/find 3D view", title="SHN Model Checker")
                return
            
            # Переключаемся на вид
            uidoc.ActiveView = view3d
            
            with Transaction(doc, "SHN: Setup Section Box") as t:
                t.Start()
                
                # Включаем Section Box
                try:
                    view3d.IsSectionBoxActive = True
                except:
                    pass
                
                # Устанавливаем bbox с padding
                if issue.issue_bbox:
                    padded = bbox_expand(issue.issue_bbox, mm_to_ft(500))  # 500mm padding
                    view3d.SetSectionBox(padded)
                
                t.Commit()
            
            # Изолируем элементы (temporary)
            ids_to_show = List[ElementId]()
            ids_to_show.Add(issue.main_id)
            
            for hid in issue.host_blocker_ids:
                ids_to_show.Add(hid)
            
            for (link_id, _) in issue.link_blockers:
                if link_id not in [x for x in ids_to_show]:
                    ids_to_show.Add(link_id)
            
            # Временная изоляция (если доступно)
            try:
                if TemporaryViewMode:
                    view3d.IsolateElementTemporary(ids_to_show)
            except:
                pass
            
            # Выделяем
            uidoc.Selection.SetElementIds(ids_to_show)
            
            # Зум
            try:
                uidoc.ShowElements(ids_to_show)
            except:
                try:
                    uidoc.ShowElements(issue.main_id)
                except:
                    pass
                    
        except Exception as ex:
            logger.error("do_3d_section error: {0}".format(ex))
    
    def find_or_create_3d_view(self):
        """Найти или создать 3D вид для проверок"""
        view_name = "SHN_Checks_3D"
        
        # Ищем существующий
        collector = FilteredElementCollector(doc).OfClass(View3D)
        for v in collector:
            try:
                if v.Name == view_name and not v.IsTemplate:
                    return v
            except:
                continue
        
        # Создаем новый
        try:
            with Transaction(doc, "Create SHN 3D View") as t:
                t.Start()
                
                # Находим ViewFamilyType для 3D
                vft_collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
                vft_3d = None
                for vft in vft_collector:
                    if vft.ViewFamily == ViewFamily.ThreeDimensional:
                        vft_3d = vft
                        break
                
                if not vft_3d:
                    t.RollBack()
                    return None
                
                new_view = View3D.CreateIsometric(doc, vft_3d.Id)
                new_view.Name = view_name
                
                t.Commit()
                return new_view
        except:
            return None

# ----------------------------
# WPF Window
# ----------------------------
XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SHN Model Checker" 
        Width="1200" Height="750"
        WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Model Info Header -->
    <Border Grid.Row="0" Background="#F0F0F0" Padding="8,6" CornerRadius="3" Margin="0,0,0,10">
      <StackPanel Orientation="Horizontal">
        <TextBlock Text="📄" FontSize="14" Margin="0,0,8,0"/>
        <TextBlock x:Name="tbModelInfo" Text="Model: Loading..." FontWeight="SemiBold" FontSize="13"/>
        <TextBlock x:Name="tbStats" Text="" FontSize="12" Foreground="#666" Margin="20,0,0,0"/>
      </StackPanel>
    </Border>

    <!-- Checks Selection -->
    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Checks:" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0"/>
      <CheckBox x:Name="cb1" Content="1) Clash vs linked models" IsChecked="True" Margin="0,0,12,0"/>
      <CheckBox x:Name="cb2" Content="2) LV service zone 800mm" IsChecked="True" Margin="0,0,12,0"/>
      <CheckBox x:Name="cb3" Content="3) Double-row corridor" IsChecked="True" Margin="0,0,12,0"/>
      <CheckBox x:Name="cb4" Content="4) Transformer clearance 800/1000" IsChecked="True" Margin="0,0,12,0"/>
      <Button x:Name="btnRun" Content="▶ Run" Width="90" Margin="20,0,0,0" Background="#4CAF50" Foreground="White" FontWeight="Bold"/>
    </StackPanel>

    <!-- Filters -->
    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <TextBox x:Name="tbFilter" Width="340" Margin="0,0,10,0" ToolTip="Filter by text (message / check / id)"/>
      <TextBlock Text="Severity:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <ComboBox x:Name="cbSeverity" Width="140" Margin="0,0,10,0"/>
      <TextBlock x:Name="txtStats" VerticalAlignment="Center" Margin="10,0,0,0" FontStyle="Italic" Foreground="#666"/>
    </StackPanel>

    <!-- Data Grid -->
    <DataGrid Grid.Row="3" x:Name="gridIssues" AutoGenerateColumns="False" IsReadOnly="True"
              SelectionMode="Single" HeadersVisibility="Column" CanUserAddRows="False"
              AlternatingRowBackground="#F9F9F9">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Status" Binding="{Binding status}" Width="90"/>
        <DataGridTextColumn Header="Severity" Binding="{Binding severity}" Width="110"/>
        <DataGridTextColumn Header="Check" Binding="{Binding check_name}" Width="240"/>
        <DataGridTextColumn Header="MainId" Binding="{Binding main_id}" Width="80"/>
        <DataGridTextColumn Header="Interfering" Binding="{Binding interfering}" Width="200"/>
        <DataGridTextColumn Header="Message" Binding="{Binding message}" Width="*"/>
      </DataGrid.Columns>
      <DataGrid.RowStyle>
        <Style TargetType="DataGridRow">
          <Style.Triggers>
            <!-- RESOLVED - зеленый -->
            <DataTrigger Binding="{Binding status}" Value="RESOLVED">
              <Setter Property="Background" Value="#C8E6C9"/>
            </DataTrigger>
            <!-- IGNORED - желтый -->
            <DataTrigger Binding="{Binding status}" Value="IGNORED">
              <Setter Property="Background" Value="#FFF9C4"/>
            </DataTrigger>
          </Style.Triggers>
        </Style>
      </DataGrid.RowStyle>
    </DataGrid>

    <!-- Action Buttons -->
    <Grid Grid.Row="4" Margin="0,10,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      
      <!-- Left side - Status buttons -->
      <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left">
        <Button x:Name="btnMarkResolved" Content="✓ Mark Resolved" Width="120" Margin="0,0,8,0" 
                Background="#4CAF50" Foreground="White" ToolTip="Mark issue as resolved"/>
        <Button x:Name="btnMarkIgnored" Content="⚠ Mark Ignored" Width="120" Margin="0,0,8,0" 
                Background="#FFC107" ToolTip="Mark issue as not requiring fix"/>
        <Button x:Name="btnMarkOpen" Content="↻ Reopen" Width="100" Margin="0,0,8,0"
                ToolTip="Mark issue as open again"/>
      </StackPanel>
      
      <!-- Right side - Navigation and action buttons -->
      <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="btnPrevious" Content="◄ Previous" Width="90" Margin="0,0,8,0"/>
        <Button x:Name="btnNext" Content="Next ►" Width="90" Margin="0,0,8,0"/>
        <Separator Width="1" Margin="8,0" Background="#CCC"/>
        <Button x:Name="btnSelect" Content="Select" Width="90" Margin="0,0,8,0"/>
        <Button x:Name="btnZoom" Content="Zoom (3D)" Width="110" Margin="0,0,8,0"/>
        <Button x:Name="btn3D" Content="3D Section" Width="110" Margin="0,0,8,0"/>
        <Button x:Name="btnClose" Content="Close" Width="90"/>
      </StackPanel>
    </Grid>

    <!-- Status Bar -->
    <Border Grid.Row="5" Background="#E8E8E8" Padding="8,4" Margin="0,10,0,0" CornerRadius="3">
      <StackPanel Orientation="Horizontal">
        <TextBlock Text="💡 Tip: " FontWeight="Bold" FontSize="11"/>
        <TextBlock Text="Double-click row to open 3D | ✓=Resolved (green) | ⚠=Ignored (yellow) | ◄►=Navigate issues" FontSize="11" Foreground="#555"/>
      </StackPanel>
    </Border>

  </Grid>
</Window>
"""

class IssueRow(object):
    def __init__(self, issue):
        self.issue = issue
        self.status = issue.status  # OPEN, RESOLVED, IGNORED
        self.severity = issue.severity
        self.check_name = issue.check_name
        self.main_id = safe_str(issue.main_id.IntegerValue)
        self.interfering = issue.interference_text()
        self.message = issue.message

class CheckerWindow(object):
    def __init__(self):
        reader = XmlReader.Create(StringReader(XAML))
        self.win = XamlReader.Load(reader)

        self.cb1 = self.win.FindName("cb1")
        self.cb2 = self.win.FindName("cb2")
        self.cb3 = self.win.FindName("cb3")
        self.cb4 = self.win.FindName("cb4")
        self.btnRun = self.win.FindName("btnRun")

        self.tbFilter = self.win.FindName("tbFilter")
        self.cbSeverity = self.win.FindName("cbSeverity")
        self.txtStats = self.win.FindName("txtStats")
        
        # Новые элементы для информации о модели
        self.tbModelInfo = self.win.FindName("tbModelInfo")
        self.tbStats = self.win.FindName("tbStats")

        self.grid = self.win.FindName("gridIssues")
        self.btnSelect = self.win.FindName("btnSelect")
        self.btnZoom = self.win.FindName("btnZoom")
        self.btn3D = self.win.FindName("btn3D")
        self.btnClose = self.win.FindName("btnClose")
        
        # Новые кнопки для статусов
        self.btnMarkResolved = self.win.FindName("btnMarkResolved")
        self.btnMarkIgnored = self.win.FindName("btnMarkIgnored")
        self.btnMarkOpen = self.win.FindName("btnMarkOpen")
        
        # Кнопки навигации
        self.btnPrevious = self.win.FindName("btnPrevious")
        self.btnNext = self.win.FindName("btnNext")

        self.nav_handler = NavHandler()
        self.ext_event = ExternalEvent.Create(self.nav_handler)

        self.rows = []
        self.filtered_rows = []  # Для навигации по отфильтрованным
        
        # Хранилище статусов
        self.storage = IssueStorage()

        self.cbSeverity.Items.Add("ALL")
        self.cbSeverity.Items.Add(Severity.CRITICAL)
        self.cbSeverity.Items.Add(Severity.WARNING)
        self.cbSeverity.Items.Add(Severity.INFO)
        self.cbSeverity.SelectedIndex = 0

        self.btnRun.Click += self.on_run
        self.btnClose.Click += self.on_close
        self.btnSelect.Click += self.on_select
        self.btnZoom.Click += self.on_zoom
        self.btn3D.Click += self.on_3d
        
        # Обработчики статусов
        self.btnMarkResolved.Click += self.on_mark_resolved
        self.btnMarkIgnored.Click += self.on_mark_ignored
        self.btnMarkOpen.Click += self.on_mark_open
        
        # Обработчики навигации
        self.btnPrevious.Click += self.on_previous
        self.btnNext.Click += self.on_next

        self.tbFilter.TextChanged += self.on_filter_changed
        self.cbSeverity.SelectionChanged += self.on_filter_changed

        # Double click row -> 3D
        try:
            self.grid.MouseDoubleClick += self.on_grid_doubleclick
        except:
            pass

        # Обновляем информацию о модели
        self.update_model_info()

        self.run_checks()

    def update_model_info(self):
        """Обновление информации о текущей модели"""
        try:
            # Имя файла модели
            doc_title = doc.Title if doc.Title else "Untitled"
            
            # Путь к файлу (если сохранен)
            doc_path = ""
            try:
                if doc.PathName:
                    import os
                    doc_path = " | " + os.path.basename(os.path.dirname(doc.PathName))
            except:
                pass
            
            # Количество связанных моделей
            links_count = len(collect_links())
            
            # Формируем строку
            info_text = "Model: {0}{1} | Links: {2}".format(
                doc_title,
                doc_path,
                links_count
            )
            
            # Обновляем UI элементы (с проверкой что они существуют)
            if self.tbModelInfo:
                self.tbModelInfo.Text = info_text
            
            # Обновляем заголовок окна
            if self.win:
                self.win.Title = "SHN Model Checker - {0}".format(doc_title)
            
        except Exception as ex:
            logger.debug("update_model_info error: {0}".format(ex))
            # Устанавливаем значение по умолчанию
            if self.tbModelInfo:
                self.tbModelInfo.Text = "Model: Unknown"

    def show(self):
        self.win.Show()

    def on_close(self, sender, args):
        try:
            self.win.Close()
        except:
            pass

    def get_selected_issue(self):
        try:
            row = self.grid.SelectedItem
            return row.issue if row else None
        except:
            return None

    def raise_nav(self, cmd_type):
        issue = self.get_selected_issue()
        if not issue:
            forms.alert("Please select a row in the report.", title="SHN Model Checker")
            return
        self.nav_handler.command = {"type": cmd_type, "issue": issue}
        self.ext_event.Raise()

    def on_select(self, sender, args):
        self.raise_nav("select")

    def on_zoom(self, sender, args):
        self.raise_nav("3d")

    def on_3d(self, sender, args):
        self.raise_nav("3d")

    def on_grid_doubleclick(self, sender, args):
        self.raise_nav("3d")
    
    # ----------------------------
    # Status Management
    # ----------------------------
    def on_mark_resolved(self, sender, args):
        """Отметить issue как исправленную"""
        row = self.grid.SelectedItem
        if not row or not row.issue:
            forms.alert("Please select an issue first.", title="Mark Resolved")
            return
        
        # Обновляем статус
        row.issue.status = IssueStatus.RESOLVED
        row.status = IssueStatus.RESOLVED
        self.storage.set_status(row.issue, IssueStatus.RESOLVED)
        
        # Обновляем отображение
        self.refresh_grid()
        logger.info("Issue marked as RESOLVED: {0}".format(self.storage.get_issue_key(row.issue)))
    
    def on_mark_ignored(self, sender, args):
        """Отметить issue как не требующую исправления"""
        row = self.grid.SelectedItem
        if not row or not row.issue:
            forms.alert("Please select an issue first.", title="Mark Ignored")
            return
        
        # Обновляем статус
        row.issue.status = IssueStatus.IGNORED
        row.status = IssueStatus.IGNORED
        self.storage.set_status(row.issue, IssueStatus.IGNORED)
        
        # Обновляем отображение
        self.refresh_grid()
        logger.info("Issue marked as IGNORED: {0}".format(self.storage.get_issue_key(row.issue)))
    
    def on_mark_open(self, sender, args):
        """Вернуть issue в статус OPEN"""
        row = self.grid.SelectedItem
        if not row or not row.issue:
            forms.alert("Please select an issue first.", title="Reopen Issue")
            return
        
        # Обновляем статус
        row.issue.status = IssueStatus.OPEN
        row.status = IssueStatus.OPEN
        self.storage.set_status(row.issue, IssueStatus.OPEN)
        
        # Обновляем отображение
        self.refresh_grid()
        logger.info("Issue reopened: {0}".format(self.storage.get_issue_key(row.issue)))
    
    def refresh_grid(self):
        """Обновить отображение grid после изменения статусов"""
        # Сохраняем текущий выбор
        selected_index = self.grid.SelectedIndex
        
        # Пересоздаем rows
        temp_rows = []
        for row in self.rows:
            temp_rows.append(IssueRow(row.issue))
        self.rows = temp_rows
        
        # Применяем фильтр
        self.apply_filter()
        
        # Восстанавливаем выбор
        if selected_index >= 0 and selected_index < len(self.filtered_rows):
            self.grid.SelectedIndex = selected_index
    
    # ----------------------------
    # Navigation
    # ----------------------------
    def on_previous(self, sender, args):
        """Переключиться на предыдущий issue"""
        if not self.filtered_rows:
            return
        
        current_index = self.grid.SelectedIndex
        if current_index <= 0:
            # Уже на первом - переходим на последний (цикл)
            self.grid.SelectedIndex = len(self.filtered_rows) - 1
        else:
            self.grid.SelectedIndex = current_index - 1
        
        # Автоматически выделяем в модели
        self.raise_nav("select")
    
    def on_next(self, sender, args):
        """Переключиться на следующий issue"""
        if not self.filtered_rows:
            return
        
        current_index = self.grid.SelectedIndex
        if current_index >= len(self.filtered_rows) - 1:
            # Уже на последнем - переходим на первый (цикл)
            self.grid.SelectedIndex = 0
        else:
            self.grid.SelectedIndex = current_index + 1
        
        # Автоматически выделяем в модели
        self.raise_nav("select")

    def on_run(self, sender, args):
        self.run_checks()

    def on_filter_changed(self, sender, args):
        self.apply_filter()

    def set_grid_rows(self, rows):
        self.grid.ItemsSource = None
        self.grid.ItemsSource = rows

    def apply_filter(self):
        text = (self.tbFilter.Text or "").strip().lower()
        sev = safe_str(self.cbSeverity.SelectedItem)

        filtered = []
        for r in self.rows:
            if sev != "ALL" and r.severity != sev:
                continue
            if text:
                blob = (r.severity + " " + r.check_name + " " + r.main_id + " " +
                        r.interfering + " " + r.message + " " + r.status).lower()
                if text not in blob:
                    continue
            filtered.append(r)

        self.filtered_rows = filtered  # Сохраняем для навигации
        self.set_grid_rows(filtered)

        # Статистика по severity
        c = sum(1 for x in self.rows if x.severity == Severity.CRITICAL)
        w = sum(1 for x in self.rows if x.severity == Severity.WARNING)
        i = sum(1 for x in self.rows if x.severity == Severity.INFO)
        
        # Статистика по status
        resolved_count = sum(1 for x in self.rows if x.status == IssueStatus.RESOLVED)
        ignored_count = sum(1 for x in self.rows if x.status == IssueStatus.IGNORED)
        open_count = sum(1 for x in self.rows if x.status == IssueStatus.OPEN)
        
        self.txtStats.Text = "Total: {0} | Critical: {1} | Warning: {2} | Info: {3} | ✓Resolved: {4} | ⚠Ignored: {5} | Open: {6}".format(
            len(self.rows), c, w, i, resolved_count, ignored_count, open_count
        )

    def run_checks(self):
        try:
            self.txtStats.Text = "Running checks..."
            if self.tbStats:
                self.tbStats.Text = "⏳ Processing..."
            self.win.Dispatcher.Invoke(lambda: None)

            geom_opt = Options()
            geom_opt.ComputeReferences = False
            geom_opt.IncludeNonVisibleObjects = False
            geom_opt.DetailLevel = ViewDetailLevel.Fine

            # Собираем оборудование
            ee = collect_family_instances_by_bic(BuiltInCategory.OST_ElectricalEquipment)
            ee_fi = [e for e in ee if isinstance(e, FamilyInstance)]

            cabinets = [fi for fi in ee_fi if is_electrical_cabinet(fi)]
            transformers = [fi for fi in ee_fi if is_transformer(fi)]

            logger.info("Found {0} cabinets, {1} transformers".format(len(cabinets), len(transformers)))

            host_obs = collect_host_obstacles()
            links = collect_links()

            logger.info("Found {0} host obstacles, {1} links".format(len(host_obs), len(links)))
            
            # Обновляем статистику объектов
            if self.tbStats:
                self.tbStats.Text = "🔍 Cabinets: {0} | Transformers: {1} | Links: {2}".format(
                    len(cabinets),
                    len(transformers),
                    len(links)
                )

            issues = []
            if self.cb1.IsChecked:
                logger.info("Running Check 1...")
                issues.extend(check_1_cabinet_clash_links(cabinets, links, geom_opt))
            if self.cb2.IsChecked:
                logger.info("Running Check 2...")
                issues.extend(check_2_lv_service_zone(cabinets, host_obs, links, geom_opt))
            if self.cb3.IsChecked:
                logger.info("Running Check 3...")
                issues.extend(check_3_double_row_corridor(cabinets, host_obs, links, geom_opt))
            if self.cb4.IsChecked:
                logger.info("Running Check 4...")
                issues.extend(check_4_transformer_clearance(transformers, host_obs, links, geom_opt))

            logger.info("Total issues found: {0}".format(len(issues)))

            # Загружаем статусы из storage
            for issue in issues:
                stored_status = self.storage.get_status(issue)
                issue.status = stored_status
            
            logger.info("Loaded statuses from storage: {0}".format(self.storage.storage_file))

            self.rows = [IssueRow(x) for x in issues]
            self.apply_filter()

        except Exception as ex:
            logger.error("run_checks error: {0}".format(ex))
            import traceback
            logger.error(traceback.format_exc())
            forms.alert("Error while running checks:\n{0}".format(ex), title="SHN Model Checker")

# ----------------------------
# ENTRY (keep one instance)
# ----------------------------
# Универсальный способ хранения состояния окна, совместимый со всеми версиями pyRevit
# Используем глобальную переменную для хранения экземпляра окна

# Проверяем существующее окно
_existing_window = None

# Пробуем разные способы получить сохраненное окно
try:
    # Способ 1: __window__ (новые версии pyRevit)
    if '__window__' in dir():
        if hasattr(__window__, 'shn_checker_window'):
            _existing_window = __window__.shn_checker_window
except:
    pass

if not _existing_window:
    try:
        # Способ 2: через __vars__ (старые версии pyRevit)
        if '__vars__' in dir():
            if hasattr(__vars__, 'shn_checker_window'):
                _existing_window = __vars__.shn_checker_window
    except:
        pass

if not _existing_window:
    try:
        # Способ 3: через __persistent__ (альтернативный способ)
        if '__persistent__' in dir():
            if hasattr(__persistent__, 'shn_checker_window'):
                _existing_window = __persistent__.shn_checker_window
    except:
        pass

# Пробуем активировать существующее окно
if _existing_window:
    try:
        if _existing_window.win and _existing_window.win.IsVisible:
            _existing_window.win.Activate()
            logger.info("Activated existing Model Checker window")
        else:
            # Окно закрыто, создаем новое
            _existing_window = None
    except:
        # Окно больше не валидно
        _existing_window = None

# Создаем новое окно если нужно
if not _existing_window:
    try:
        w = CheckerWindow()
        
        # Сохраняем окно всеми доступными способами
        try:
            if '__window__' in dir():
                __window__.shn_checker_window = w
        except:
            pass
        
        try:
            if '__vars__' in dir():
                __vars__.shn_checker_window = w
        except:
            pass
        
        try:
            if '__persistent__' in dir():
                __persistent__.shn_checker_window = w
        except:
            pass
        
        w.show()
        logger.info("SHN Model Checker window opened")
    except Exception as ex:
        logger.error(ex)
        import traceback
        logger.error(traceback.format_exc())
        forms.alert("SHN Model Checker crashed:\n{0}".format(ex), title="SHN Model Checker")
