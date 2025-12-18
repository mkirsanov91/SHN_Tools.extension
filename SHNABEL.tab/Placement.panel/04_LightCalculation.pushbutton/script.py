# -*- coding: utf-8 -*-
"""
Export room boundaries from selected Revit link to DXF per Level
for automatic room detection in DIALux evo (free).

For each Level in the linked architectural model:
    - collect all Rooms on that level
    - take their boundary loops
    - transform to host coordinates, convert to meters
    - write a separate DXF (R12) with closed POLYLINE entities on layer "ROOMS"

In DIALux evo:
    - For each floor, import corresponding DXF as plan
    - Use "Create rooms from CAD polylines" to generate rooms automatically
"""
__title__ = 'Export\nRooms DXF'
__author__ = 'SHNABEL digital'

import clr
import System
import os

from pyrevit import revit, forms

clr.AddReference('RevitAPI')
import Autodesk.Revit.DB as DB

doc = revit.doc
FT_TO_M = 0.3048


# -----------------------------------------------------------------------------
# 1. Выбор архитектурного линка из списка
# -----------------------------------------------------------------------------
class LinkItem(object):
    def __init__(self, inst, link_doc):
        self.inst = inst
        self.link_doc = link_doc
        self._display = u"{}  (instance: {})".format(link_doc.Title, inst.Name)

    @property
    def display(self):
        return self._display


link_instances = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.RevitLinkInstance)
    .WhereElementIsNotElementType()
)

link_items = []
for li in link_instances:
    try:
        ld = li.GetLinkDocument()
        if ld:
            link_items.append(LinkItem(li, ld))
    except:
        pass

if not link_items:
    forms.alert(
        "No Revit links found in current model.\n"
        "This tool expects an architectural model linked.",
        title="Export Rooms DXF",
        warn_icon=True
    )
    raise SystemExit

sel = forms.SelectFromList.show(
    link_items,
    title="Select architectural link to export rooms as DXF (per level)",
    multiselect=False,
    name_attr='display'
)

if not sel:
    raise SystemExit

link_inst = sel.inst
link_doc = sel.link_doc
link_tr = link_inst.GetTransform()  # transform from link to host coords


# -----------------------------------------------------------------------------
# 2. Сбор комнат по уровням
# -----------------------------------------------------------------------------
opts = DB.SpatialElementBoundaryOptions()
opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish

rooms = DB.FilteredElementCollector(link_doc)\
    .OfCategory(DB.BuiltInCategory.OST_Rooms)\
    .WhereElementIsNotElementType()\
    .ToElements()

# словарь: level_name -> [poly_1, poly_2, ...]
# poly = [(x_m, y_m), ...]  (замкнутый контур, последняя точка = первая)
level_polygons = {}

def norm_text(s):
    if not s:
        return u""
    return s.replace(u'\ufeff', u'').replace(u'\u200f', u'').strip()

for room in rooms:
    try:
        if room.Area <= 0 or not room.Location:
            continue

        level = room.Level
        lvl_name = norm_text(level.Name if level else "NoLevel")

        boundaries = room.GetBoundarySegments(opts)
        if not boundaries:
            continue

        # обычно первая петля — внешний контур
        loop = boundaries[0]
        pts = []
        for seg in loop:
            curve = seg.GetCurve()
            p = curve.GetEndPoint(0)
            # координаты линка -> в хост
            p_host = link_tr.OfPoint(p)
            x_m = p_host.X * FT_TO_M
            y_m = p_host.Y * FT_TO_M
            pts.append((x_m, y_m))

        if len(pts) < 3:
            continue

        # закрываем полигон
        if pts[0] != pts[-1]:
            pts.append(pts[0])

        if lvl_name not in level_polygons:
            level_polygons[lvl_name] = []
        level_polygons[lvl_name].append(pts)

    except Exception as e:
        print("Error processing room {} in link {}: {}".format(
            room.Id, link_doc.Title, e)
        )

if not level_polygons:
    forms.alert(
        "No valid room boundaries found in link:\n{}".format(link_doc.Title),
        title="Export Rooms DXF",
        warn_icon=True
    )
    raise SystemExit


# -----------------------------------------------------------------------------
# 3. Выбор папки
# -----------------------------------------------------------------------------
folder = forms.pick_folder(title="Select folder to save DXF files for DIALux")
if not folder:
    raise SystemExit

proj_name = doc.ProjectInformation.Name or doc.Title or "Revit_Project"
if ".rvt" in proj_name.lower():
    proj_name = proj_name.replace(".rvt", "").replace(".RVT", "")

invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
for ch in invalid:
    proj_name = proj_name.replace(ch, "_")


def safe_name(s):
    if not s:
        return "NoLevel"
    for ch in invalid:
        s = s.replace(ch, "_")
    return s


# -----------------------------------------------------------------------------
# 4. Функция записи DXF R12 (AC1009) с POLYLINE/VERTEX
# -----------------------------------------------------------------------------
def write_dxf(polys, path):
    lines = []

    # HEADER – минимальный
    lines.extend([
        "0", "SECTION",
        "2", "HEADER",
        "9", "$ACADVER",
        "1", "AC1009",     # R12 – максимально совместимая версия
        "0", "ENDSEC",
    ])

    # TABLES – определяем слой ROOMS
    lines.extend([
        "0", "SECTION",
        "2", "TABLES",
        "0", "TABLE",
        "2", "LAYER",
        "70", "1",       # number of entries
        "0", "LAYER",
        "2", "ROOMS",    # layer name
        "70", "0",
        "62", "7",       # color (7 = white)
        "6", "CONTINUOUS",
        "0", "ENDTAB",
        "0", "ENDSEC",
    ])

    # ENTITIES – наши полигоны
    lines.extend([
        "0", "SECTION",
        "2", "ENTITIES",
    ])

    for poly in polys:
        n = len(poly)
        if n < 3:
            continue

        # Нормальная polyline с вершинами
        lines.extend([
            "0", "POLYLINE",
            "8", "ROOMS",   # layer
            "66", "1",      # vertices follow
            "70", "1",      # closed polyline
        ])

        for (x, y) in poly:
            lines.extend([
                "0", "VERTEX",
                "8", "ROOMS",
                "10", "{:.6f}".format(x),
                "20", "{:.6f}".format(y),
            ])

        lines.extend([
            "0", "SEQEND",
        ])

    # END ENTITIES
    lines.extend([
        "0", "ENDSEC",
        "0", "EOF",
    ])

    with open(path, "w") as f:
        f.write("\n".join(lines))


# -----------------------------------------------------------------------------
# 5. Для каждого уровня – свой DXF, с локальной нормализацией координат
# -----------------------------------------------------------------------------
created_files = []

for lvl_name, polys in level_polygons.items():
    # нормализуем координаты в пределах уровня, чтобы план был рядом с (0,0)
    min_x = min(pt[0] for poly in polys for pt in poly)
    min_y = min(pt[1] for poly in polys for pt in poly)

    norm_polys = []
    for poly in polys:
        norm = [(x - min_x, y - min_y) for (x, y) in poly]
        norm_polys.append(norm)

    lvl_safe = safe_name(lvl_name)
    filename = u"{}_{}_RoomsForDialux.dxf".format(proj_name, lvl_safe)
    filepath = os.path.join(folder, filename)

    write_dxf(norm_polys, filepath)
    created_files.append(filepath)


# -----------------------------------------------------------------------------
# 6. Сообщение пользователю
# -----------------------------------------------------------------------------
msg = "DXF export finished.\n\nCreated files:\n"
for fp in created_files:
    msg += "  - {}\n".format(fp)

msg += (
    "\nIn DIALux evo for each level:\n"
    "  1) Import corresponding DXF as plan\n"
    "  2) Use 'Create rooms from CAD polylines' to generate rooms automatically."
)

forms.alert(msg, title="Export Rooms DXF", warn_icon=False)
