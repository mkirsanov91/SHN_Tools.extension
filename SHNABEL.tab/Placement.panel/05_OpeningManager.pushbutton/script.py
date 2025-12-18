# -*- coding: utf-8 -*-
# SHN Opening Browser (v0.2) - pyRevit / IronPython
# FIXED: Manager window now properly displays and navigates openings
# - Place/Update face-based Generic Model openings on LINKED walls
# - Openings are oriented 90° to wall (Depth along wall normal)
# - Cluster nearby penetrations (tray/pipe packages) into one opening
# - Audit: Missing / NeedResize / Empty

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("RevitAPIUI")

from System.Windows.Markup import XamlReader
from System.Windows.Interop import WindowInteropHelper
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
uiapp = uidoc.Application

# -----------------------------
# Defaults / Parameter name fallbacks
# -----------------------------
DEFAULT_FAMILY_NAME = "SHN_Openings_GEN_Square FaceBased"

WIDTH_NAMES  = ["Width"]
HEIGHT_NAMES = ["Highet", "Hight", "Height"]
DEPTH_NAMES  = ["Depth"]

FROM_LINK_NAMES    = ["From Link", "SHN_FromLink"]
ID_IN_LINK_NAMES   = ["ID in Link", "SHN_IDinLink"]
APPROVED_NAMES     = ["Approved", "SHN_Approved"]
CHANGED_NAMES      = ["Changed", "SHN_Changed"]
NEW_NAMES          = ["New", "SHN_New"]


# -----------------------------
# Units
# -----------------------------
def mm_to_ft(mm):
    return float(mm) / 304.8

def ft_to_mm(ft):
    return float(ft) * 304.8


# -----------------------------
# Safe parameter helpers
# -----------------------------
def lookup_param_any(elem, names):
    if not elem:
        return None
    for n in names:
        p = elem.LookupParameter(n)
        if p:
            return p
    return None

def get_double_param(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p and p.HasValue:
            return p.AsDouble()
    except:
        pass
    return None

def set_param(p, value):
    if not p:
        return False
    try:
        st = p.StorageType
        if st == DB.StorageType.Double:
            p.Set(float(value))
        elif st == DB.StorageType.Integer:
            p.Set(int(value))
        elif st == DB.StorageType.String:
            p.Set(str(value))
        elif st == DB.StorageType.ElementId:
            if isinstance(value, DB.ElementId):
                p.Set(value)
            else:
                p.Set(DB.ElementId(int(value)))
        else:
            p.Set(str(value))
        return True
    except:
        return False

def get_param_as_int(elem, names, default_val=None):
    p = lookup_param_any(elem, names)
    if not p:
        return default_val
    try:
        if p.StorageType == DB.StorageType.Integer:
            return p.AsInteger()
        if p.StorageType == DB.StorageType.Double:
            return int(round(p.AsDouble()))
        if p.StorageType == DB.StorageType.String:
            s = p.AsString()
            if s is None:
                return default_val
            return int(s)
    except:
        return default_val
    return default_val

def is_param_checked(elem, names):
    p = lookup_param_any(elem, names)
    if not p:
        return False
    try:
        if p.StorageType == DB.StorageType.Integer:
            return p.AsInteger() == 1
        if p.StorageType == DB.StorageType.Double:
            return abs(p.AsDouble() - 1.0) < 0.0001
        if p.StorageType == DB.StorageType.String:
            s = p.AsString()
            return (s or "").strip().lower() in ["1", "true", "yes"]
    except:
        return False
    return False

def ensure_symbol_active(sym):
    if not sym:
        return
    try:
        if sym.IsActive:
            return
        if not doc.IsModifiable:
            t = DB.Transaction(doc, "SHN: Activate opening type")
            t.Start()
            sym.Activate()
            doc.Regenerate()
            t.Commit()
        else:
            sym.Activate()
            doc.Regenerate()
    except:
        pass


# -----------------------------
# Geometry helpers
# -----------------------------
def norm(xyz):
    try:
        if xyz and xyz.GetLength() > 1e-9:
            return xyz.Normalize()
    except:
        pass
    return None

def aabb_from_bbox(bb, trf=None):
    if not bb:
        return None
    pts = []
    mn = bb.Min
    mx = bb.Max
    pts.append(DB.XYZ(mn.X, mn.Y, mn.Z))
    pts.append(DB.XYZ(mn.X, mn.Y, mx.Z))
    pts.append(DB.XYZ(mn.X, mx.Y, mn.Z))
    pts.append(DB.XYZ(mn.X, mx.Y, mx.Z))
    pts.append(DB.XYZ(mx.X, mn.Y, mn.Z))
    pts.append(DB.XYZ(mx.X, mn.Y, mx.Z))
    pts.append(DB.XYZ(mx.X, mx.Y, mn.Z))
    pts.append(DB.XYZ(mx.X, mx.Y, mx.Z))

    if trf:
        pts = [trf.OfPoint(p) for p in pts]

    minx = min([p.X for p in pts]); miny = min([p.Y for p in pts]); minz = min([p.Z for p in pts])
    maxx = max([p.X for p in pts]); maxy = max([p.Y for p in pts]); maxz = max([p.Z for p in pts])
    return (DB.XYZ(minx, miny, minz), DB.XYZ(maxx, maxy, maxz))

def aabb_intersects(a_min, a_max, b_min, b_max, tol=0.0):
    if not a_min or not a_max or not b_min or not b_max:
        return False
    if (a_max.X + tol) < b_min.X or (b_max.X + tol) < a_min.X:
        return False
    if (a_max.Y + tol) < b_min.Y or (b_max.Y + tol) < a_min.Y:
        return False
    if (a_max.Z + tol) < b_min.Z or (b_max.Z + tol) < a_min.Z:
        return False
    return True

def iter_solids(geom_elem):
    if not geom_elem:
        return
    for g in geom_elem:
        if isinstance(g, DB.Solid):
            if g.Volume > 1e-9:
                yield g
        elif isinstance(g, DB.GeometryInstance):
            try:
                inst = g.GetInstanceGeometry()
                for s in iter_solids(inst):
                    yield s
            except:
                pass

def planar_faces_from_solid(solid):
    faces = []
    try:
        for f in solid.Faces:
            if isinstance(f, DB.PlanarFace):
                n = f.FaceNormal
                if abs(n.Z) < 0.3:
                    faces.append(f)
    except:
        pass
    return faces

def curve_direction(curve):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return norm(p1 - p0)
    except:
        return None

def project_point_to_plane(p, plane_origin, plane_normal):
    v = p - plane_origin
    d = v.DotProduct(plane_normal)
    return p - (plane_normal.Multiply(d))


# -----------------------------
# Rect / clustering
# -----------------------------
class RectUV(object):
    def __init__(self, umin, umax, vmin, vmax, src_ids=None):
        self.umin = umin; self.umax = umax
        self.vmin = vmin; self.vmax = vmax
        self.src_ids = set(src_ids) if src_ids else set()

    def center(self):
        return ((self.umin + self.umax) * 0.5, (self.vmin + self.vmax) * 0.5)

    def width(self):
        return (self.umax - self.umin)

    def height(self):
        return (self.vmax - self.vmin)

    def merge(self, other):
        self.umin = min(self.umin, other.umin)
        self.umax = max(self.umax, other.umax)
        self.vmin = min(self.vmin, other.vmin)
        self.vmax = max(self.vmax, other.vmax)
        self.src_ids |= other.src_ids

def rect_intersects_with_gap(a, b, gap):
    if (a.umax + gap) < b.umin: return False
    if (b.umax + gap) < a.umin: return False
    if (a.vmax + gap) < b.vmin: return False
    if (b.vmax + gap) < a.vmin: return False
    return True

def cluster_rects(rects, gap):
    clusters = []
    for r in rects:
        merged = False
        for c in clusters:
            if rect_intersects_with_gap(c, r, gap):
                c.merge(r)
                merged = True
                break
        if not merged:
            clusters.append(RectUV(r.umin, r.umax, r.vmin, r.vmax, r.src_ids))

    changed = True
    while changed:
        changed = False
        out = []
        while clusters:
            c = clusters.pop(0)
            i = 0
            while i < len(clusters):
                if rect_intersects_with_gap(c, clusters[i], gap):
                    c.merge(clusters[i])
                    clusters.pop(i)
                    changed = True
                else:
                    i += 1
            out.append(c)
        clusters = out
    return clusters


# -----------------------------
# Collectors / UI
# -----------------------------
def get_symbol_names(sym):
    fam_name = None
    type_name = None
    try:
        pf = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if pf: fam_name = pf.AsString()
    except:
        fam_name = None
    try:
        pt = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if pt: type_name = pt.AsString()
    except:
        type_name = None

    if not fam_name:
        try:
            fam_name = sym.Family.Name
        except:
            fam_name = "Family"

    if not type_name:
        try:
            type_name = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        except:
            type_name = "Type"

    return fam_name, type_name

def pick_opening_symbol():
    symbols = []
    fec = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
    for s in fec:
        fam = None
        try:
            fam = s.Family
        except:
            continue
        if not fam:
            continue
        cat = fam.FamilyCategory
        if cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_GenericModel):
            symbols.append(s)

    if not symbols:
        forms.alert("No Generic Model family symbols were found in this model.", exitscript=True)

    items = []
    default_item = None
    for s in symbols:
        fam_name, type_name = get_symbol_names(s)
        nm = "{} : {}".format(fam_name, type_name)
        items.append(nm)
        if fam_name == DEFAULT_FAMILY_NAME and default_item is None:
            default_item = nm

    chosen = forms.SelectFromList.show(items,
                                       title="Select opening family type (Generic Model / Face-Based)",
                                       multiselect=False,
                                       default=default_item,
                                       button_name="OK")
    if not chosen:
        script.exit()

    for s in symbols:
        fam_name, type_name = get_symbol_names(s)
        nm = "{} : {}".format(fam_name, type_name)
        if nm == chosen:
            return s
    return symbols[0]

def safe_link_name(li):
    try:
        p = li.get_Parameter(DB.BuiltInParameter.RVT_LINK_INSTANCE_NAME)
        if p and p.HasValue:
            return p.AsString()
    except:
        pass
    try:
        return li.Name
    except:
        return None

def link_display(li):
    title = None
    try:
        ld = li.GetLinkDocument()
        if ld:
            title = ld.Title
    except:
        title = None

    nm = safe_link_name(li)
    base = title or nm or "Link"
    if base and (".rvt" not in base.lower()):
        base = base + ".rvt"
    try:
        return "{}  | InstId {}".format(base, li.Id.IntegerValue)
    except:
        return "{}  | InstId ?".format(base)

def pick_link_instances():
    links = []
    fec = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
    for li in fec:
        try:
            if li.GetLinkDocument():
                links.append(li)
        except:
            pass
    if not links:
        forms.alert("No loaded Revit links were found (RevitLinkInstance).", exitscript=True)

    items = []
    by_item = {}
    for li in links:
        disp = link_display(li)
        if disp in by_item:
            disp = "{}  (Unique {})".format(disp, li.Id.IntegerValue)
        items.append(disp)
        by_item[disp] = li

    chosen = forms.SelectFromList.show(items,
                                       title="Select link model(s) that contain walls",
                                       multiselect=True,
                                       button_name="OK")
    if not chosen:
        script.exit()
    return [by_item[x] for x in chosen]

def ask_float(prompt, default_val):
    s = forms.ask_for_string(default=str(default_val), prompt=prompt, title="SHN Opening Browser")
    if s is None:
        script.exit()
    try:
        return float(str(s).replace(",", "."))
    except:
        forms.alert("Could not parse number: {}".format(s), exitscript=True)

def get_mep_elements():
    cats = List[DB.BuiltInCategory]()
    cats.Add(DB.BuiltInCategory.OST_CableTray)
    cats.Add(DB.BuiltInCategory.OST_Conduit)
    flt = DB.ElementMulticategoryFilter(cats)
    els = DB.FilteredElementCollector(doc).WherePasses(flt).WhereElementIsNotElementType().ToElements()
    return [e for e in els]

def collect_existing_openings(opening_family_id):
    fam_int = None
    try:
        fam_int = opening_family_id.IntegerValue
    except:
        try:
            fam_int = int(opening_family_id)
        except:
            fam_int = None
    if fam_int is None:
        return []

    res = []
    fec = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).WhereElementIsNotElementType()
    for e in fec:
        try:
            sym = e.Symbol
            if sym and sym.Family and sym.Family.Id and sym.Family.Id.IntegerValue == fam_int:
                res.append(e)
        except:
            pass
    return res


# -----------------------------
# Wall data cache
# -----------------------------
def build_walls_cache(link_instances):
    opt = DB.Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    opt.DetailLevel = DB.ViewDetailLevel.Fine

    walls_cache = []
    for li in link_instances:
        ld = li.GetLinkDocument()
        if not ld:
            continue
        trf = li.GetTransform()
        inv = trf.Inverse

        try:
            walls = DB.FilteredElementCollector(ld).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
        except:
            walls = []

        for w in walls:
            try:
                loc = w.Location
                if not isinstance(loc, DB.LocationCurve):
                    continue
                c = loc.Curve
                d_link = curve_direction(c)
                if not d_link:
                    continue

                U = norm(trf.OfVector(d_link))
                if not U:
                    continue
                V = DB.XYZ.BasisZ
                N = norm(U.CrossProduct(V))
                if not N:
                    N = norm(V.CrossProduct(U))
                if not N:
                    continue

                bb_link = w.get_BoundingBox(None)
                aabb = aabb_from_bbox(bb_link, trf)
                if not aabb:
                    continue

                try:
                    thickness = float(w.Width)
                except:
                    thickness = mm_to_ft(200.0)

                geom = w.get_Geometry(opt)
                faces_link = []
                for s in iter_solids(geom):
                    pf = planar_faces_from_solid(s)
                    for f in pf:
                        try:
                            if f and f.Reference:
                                faces_link.append(f)
                        except:
                            pass

                if not faces_link:
                    continue

                best_face = None
                best_dot = -1.0
                best_n_host = None
                best_o_host = None

                for f in faces_link:
                    try:
                        n_host = norm(trf.OfVector(f.FaceNormal))
                        if not n_host:
                            continue
                        dot = abs(n_host.DotProduct(N))
                        if dot > best_dot:
                            best_dot = dot
                            best_face = f
                            best_n_host = n_host
                            best_o_host = trf.OfPoint(f.Origin)
                    except:
                        pass

                if not best_face or not best_n_host or not best_o_host:
                    continue

                try:
                    bbuv = best_face.GetBoundingBox()
                except:
                    continue

                try:
                    o_link = c.GetEndPoint(0)
                except:
                    o_link = c.Evaluate(0.0, True)
                O = trf.OfPoint(o_link)

                walls_cache.append({
                    "link_inst": li,
                    "link_id": li.Id.IntegerValue,
                    "wall": w,
                    "wall_id": w.Id.IntegerValue,

                    "aabb_min": aabb[0],
                    "aabb_max": aabb[1],

                    "inv_trf": inv,
                    "face_bbuv": bbuv,

                    "face_link": best_face,
                    "face_n_host": best_n_host,
                    "face_o_host": best_o_host,

                    "O": O, "U": U, "V": V, "N": N,
                    "thickness": thickness
                })
            except:
                pass

    return walls_cache

def make_link_face_reference(link_inst, face_ref):
    try:
        return face_ref.CreateLinkReference(link_inst)
    except:
        try:
            return DB.Reference.CreateLinkReference(link_inst, face_ref)
        except:
            return None


# -----------------------------
# Core computation
# -----------------------------
def element_required_wh(elem, clearance_ft):
    catid = elem.Category.Id.IntegerValue if elem.Category else None

    if catid == int(DB.BuiltInCategory.OST_CableTray):
        w = get_double_param(elem, DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        h = get_double_param(elem, DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
        if w is None: w = mm_to_ft(200.0)
        if h is None: h = mm_to_ft(100.0)
        return (w + clearance_ft, h + clearance_ft)

    if catid == int(DB.BuiltInCategory.OST_Conduit):
        d = get_double_param(elem, DB.BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
        if d is None:
            d = mm_to_ft(25.0)
        side = d + clearance_ft
        return (side, side)

    return None

def line_plane_intersection(p0, p1, plane_origin, plane_normal):
    v = p1 - p0
    denom = v.DotProduct(plane_normal)
    if abs(denom) < 1e-9:
        return (None, None)
    t = (plane_origin - p0).DotProduct(plane_normal) / denom
    hp = p0 + v.Multiply(t)
    return (hp, t)

def compute_penetrations(mep_elems, walls_cache, clearance_ft, bbox_tol_ft):
    out = {}
    tol = mm_to_ft(5.0)

    for e in mep_elems:
        try:
            loc = e.Location
            if not isinstance(loc, DB.LocationCurve):
                continue
            curve = loc.Curve
            if not curve:
                continue

            line = curve if isinstance(curve, DB.Line) else None
            if not line:
                continue

            p0 = line.GetEndPoint(0)
            p1 = line.GetEndPoint(1)

            bb = e.get_BoundingBox(None)
            if not bb:
                continue
            e_min, e_max = bb.Min, bb.Max

            wh = element_required_wh(e, clearance_ft)
            if not wh:
                continue
            base_w, base_h = wh[0], wh[1]

            D = norm(p1 - p0)
            if not D:
                continue

            for wd in walls_cache:
                if not aabb_intersects(e_min, e_max, wd["aabb_min"], wd["aabb_max"], tol=bbox_tol_ft):
                    continue

                O = wd["O"]; U = wd["U"]; V = wd["V"]; N = wd["N"]
                t_wall = wd["thickness"]
                half = t_wall * 0.5

                d0 = (p0 - O).DotProduct(N)
                d1 = (p1 - O).DotProduct(N)

                strict_cross = ((d0 >  half + tol and d1 < -half - tol) or
                                (d1 >  half + tol and d0 < -half - tol))
                if not strict_cross:
                    continue

                hp, tseg = line_plane_intersection(p0, p1, O, N)
                if hp is None:
                    continue
                if tseg < -0.01 or tseg > 1.01:
                    continue

                face_link = wd["face_link"]
                inv = wd["inv_trf"]
                bbuv = wd["face_bbuv"]

                hp_link = inv.OfPoint(hp)
                proj = face_link.Project(hp_link)
                if proj is None:
                    continue

                uv = proj.UVPoint
                if uv.U < (bbuv.Min.U - 1e-6) or uv.U > (bbuv.Max.U + 1e-6):
                    continue
                if uv.V < (bbuv.Min.V - 1e-6) or uv.V > (bbuv.Max.V + 1e-6):
                    continue

                c = abs(D.DotProduct(N))
                if c < 1e-3:
                    continue
                du = (t_wall * abs(D.DotProduct(U))) / c
                dv = (t_wall * abs(D.DotProduct(V))) / c

                req_w = base_w + du
                req_h = base_h + dv

                vOP = hp - O
                u = vOP.DotProduct(U)
                v = vOP.DotProduct(V)

                r = RectUV(u - req_w * 0.5, u + req_w * 0.5,
                           v - req_h * 0.5, v + req_h * 0.5,
                           src_ids=[e.Id.IntegerValue])

                key = (wd["link_id"], wd["wall_id"])
                if key not in out:
                    out[key] = []
                out[key].append(r)

        except:
            pass

    return out

def walldata_by_key(walls_cache):
    d = {}
    for wd in walls_cache:
        key = (wd["link_id"], wd["wall_id"])
        d[key] = wd
    return d

def opening_rect_uv(opening, wd):
    try:
        lp = opening.Location
        if not isinstance(lp, DB.LocationPoint):
            return None
        p = lp.Point
        O = wd["O"]; U = wd["U"]; V = wd["V"]
        vOP = p - O
        u = vOP.DotProduct(U)
        v = vOP.DotProduct(V)

        pw = lookup_param_any(opening, WIDTH_NAMES)
        ph = lookup_param_any(opening, HEIGHT_NAMES)
        if not pw or not ph:
            return None
        w = pw.AsDouble()
        h = ph.AsDouble()
        return RectUV(u - w * 0.5, u + w * 0.5, v - h * 0.5, v + h * 0.5, src_ids=None)
    except:
        return None

def find_best_opening_for_request(req_center_uv, openings, wd, search_tol_ft):
    best = None
    best_d = 1e99
    cu, cv = req_center_uv
    for op in openings:
        r = opening_rect_uv(op, wd)
        if not r:
            continue
        ou, ov = r.center()
        d = ((ou - cu) * (ou - cu) + (ov - cv) * (ov - cv)) ** 0.5
        if d < best_d:
            best_d = d
            best = op
    if best and best_d <= search_tol_ft:
        return best
    return None


# -----------------------------
# Placement / update
# -----------------------------
def move_instance_to_point(inst, target_point):
    try:
        lp = inst.Location
        if isinstance(lp, DB.LocationPoint):
            cur = lp.Point
            delta = target_point - cur
            if delta.GetLength() > 1e-6:
                DB.ElementTransformUtils.MoveElement(doc, inst.Id, delta)
            return True
    except:
        pass
    return False

def set_opening_params(inst, width_ft, height_ft, depth_ft, link_id_int, wall_id_int, mark_new=False, mark_changed=False):
    pw = lookup_param_any(inst, WIDTH_NAMES)
    ph = lookup_param_any(inst, HEIGHT_NAMES)
    pd = lookup_param_any(inst, DEPTH_NAMES)

    set_param(pw, width_ft)
    set_param(ph, height_ft)
    set_param(pd, depth_ft)

    set_param(lookup_param_any(inst, FROM_LINK_NAMES), link_id_int)
    set_param(lookup_param_any(inst, ID_IN_LINK_NAMES), wall_id_int)

    if mark_new:
        p = lookup_param_any(inst, NEW_NAMES)
        if p: set_param(p, 1)
    if mark_changed:
        p = lookup_param_any(inst, CHANGED_NAMES)
        if p: set_param(p, 1)

def place_or_update_openings(opening_symbol, walls_cache, clustered_by_wall, depth_extra_ft, search_tol_ft, mode_create_update):
    wd_by_key = walldata_by_key(walls_cache)
    existing_openings = collect_existing_openings(opening_symbol.Family.Id)

    openings_by_wall = {}
    for op in existing_openings:
        lk = get_param_as_int(op, FROM_LINK_NAMES, default_val=None)
        wid = get_param_as_int(op, ID_IN_LINK_NAMES, default_val=None)
        if lk is None or wid is None:
            continue
        key = (lk, wid)
        if key not in openings_by_wall:
            openings_by_wall[key] = []
        openings_by_wall[key].append(op)

    created = []
    updated = []
    missing_requests = []
    need_resize = []
    matched_openings = set()

    t = DB.Transaction(doc, "SHN Openings - Create/Update" if mode_create_update else "SHN Openings - Audit")
    if mode_create_update:
        t.Start()
        ensure_symbol_active(opening_symbol)

    for key, clusters in clustered_by_wall.items():
        wd = wd_by_key.get(key, None)
        if not wd:
            continue

        cand = openings_by_wall.get(key, [])

        for cl in clusters:
            req_w = cl.width()
            req_h = cl.height()
            req_u, req_v = cl.center()

            O = wd["O"]; U = wd["U"]; V = wd["V"]
            p_center = O + U.Multiply(req_u) + V.Multiply(req_v)

            face_o = wd["face_o_host"]
            face_n = wd["face_n_host"]
            p_on_face = project_point_to_plane(p_center, face_o, face_n)

            depth_ft = wd["thickness"] + depth_extra_ft

            op = find_best_opening_for_request((req_u, req_v), cand, wd, search_tol_ft)

            if not op:
                missing_requests.append((key, cl))
                if mode_create_update:
                    link_inst = wd["link_inst"]
                    face_link = wd["face_link"]
                    link_face_ref = make_link_face_reference(link_inst, face_link.Reference)
                    if not link_face_ref:
                        continue

                    ref_dir = U
                    try:
                        new_inst = doc.Create.NewFamilyInstance(link_face_ref, p_on_face, ref_dir, opening_symbol)
                        set_opening_params(new_inst, req_w, req_h, depth_ft, key[0], key[1], mark_new=True, mark_changed=False)
                        created.append(new_inst.Id.IntegerValue)
                        matched_openings.add(new_inst.Id.IntegerValue)

                        cand.append(new_inst)
                        if key not in openings_by_wall:
                            openings_by_wall[key] = []
                        openings_by_wall[key].append(new_inst)
                    except:
                        pass
                continue

            matched_openings.add(op.Id.IntegerValue)

            pw = lookup_param_any(op, WIDTH_NAMES)
            ph = lookup_param_any(op, HEIGHT_NAMES)
            if not pw or not ph:
                continue
            cur_w = pw.AsDouble()
            cur_h = ph.AsDouble()

            need = (req_w > (cur_w + mm_to_ft(1.0))) or (req_h > (cur_h + mm_to_ft(1.0)))
            if need:
                need_resize.append(op.Id.IntegerValue)

            if mode_create_update:
                move_instance_to_point(op, p_on_face)
                new_w = max(cur_w, req_w)
                new_h = max(cur_h, req_h)
                set_opening_params(op, new_w, new_h, depth_ft, key[0], key[1], mark_new=False, mark_changed=need)
                updated.append(op.Id.IntegerValue)

    if mode_create_update:
        t.Commit()

    empty = []
    for op in existing_openings:
        oid = op.Id.IntegerValue
        if oid in matched_openings:
            continue
        if is_param_checked(op, APPROVED_NAMES):
            continue
        empty.append(oid)

    return (created, updated, missing_requests, need_resize, empty)


# -----------------------------
# Opening Manager (MODELLESS) - FIXED VERSION
# -----------------------------
_MANAGER_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SHN Opening Manager"
        Height="720" Width="820"
        WindowStartupLocation="CenterScreen"
        Topmost="False"
        ShowInTaskbar="False">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock x:Name="TxtHeader" Grid.Row="0" FontSize="14" FontWeight="Bold" Margin="0,0,0,6" />
    <TextBlock x:Name="TxtInfo" Grid.Row="1" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" />

    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Source:" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <ComboBox x:Name="CmbSource" Width="320">
        <ComboBoxItem Content="This Run (Created/Updated)" IsSelected="True"/>
        <ComboBoxItem Content="All in Model (Selected family)"/>
      </ComboBox>
      <TextBlock x:Name="TxtCounts" VerticalAlignment="Center" Margin="12,0,0,0" Foreground="Gray"/>
    </StackPanel>

    <GroupBox Grid.Row="3" Header="Openings (double-click = Isolate ONE)" Margin="0,0,0,10">
      <Grid Margin="8">
        <Grid.RowDefinitions>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <ListBox x:Name="LstOpenings" Grid.Row="0" />

        <TextBlock Grid.Row="1" Margin="0,8,0,0" FontSize="11" Foreground="Gray"
                   Text="Tip: Revit stays active. Click in Revit view to orbit/pan. Use Isolate to focus on an opening." />
      </Grid>
    </GroupBox>

    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Left">
      <Button x:Name="BtnPrev" Content="Prev" Width="70" Margin="0,0,6,0"/>
      <Button x:Name="BtnNext" Content="Next" Width="70" Margin="0,0,6,0"/>
      <Button x:Name="BtnRefresh" Content="Refresh" Width="90" Margin="0,0,6,0"/>
    </StackPanel>

    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnSelect" Content="Select" Width="70" Margin="0,0,6,0"/>
      <Button x:Name="BtnShow" Content="Show" Width="70" Margin="0,0,6,0"/>
      <Button x:Name="BtnIsolate" Content="Isolate" Width="80" Margin="0,0,6,0"/>
      <Button x:Name="BtnClearIso" Content="Clear isolate" Width="95" Margin="0,0,6,0"/>
      <Button x:Name="BtnDelete" Content="Delete" Width="70" Margin="0,0,6,0"/>
      <Button x:Name="BtnClose" Content="Close" Width="70"/>
    </StackPanel>

    <TextBlock x:Name="TxtFooter" Grid.Row="5" FontSize="11" Foreground="Gray" Margin="0,8,0,0"
               Text="Modeless window: Revit remains active. Actions run through ExternalEvent."/>
  </Grid>
</Window>
"""

def _load_window_from_xaml(xaml_text):
    return XamlReader.Parse(xaml_text)

def get_level_name(inst):
    """Best-effort level name. Always returns non-empty."""
    try:
        lid = inst.LevelId
        if lid and lid.IntegerValue > 0:
            lvl = doc.GetElement(lid)
            if lvl and (lvl.Name or "").strip():
                return (lvl.Name or "").strip()
    except:
        pass

    bips = []
    try: bips.append(DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
    except: pass
    try: bips.append(DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM)
    except: pass
    try: bips.append(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
    except: pass

    for bip in bips:
        try:
            p = inst.get_Parameter(bip)
            if p and p.HasValue:
                lid2 = p.AsElementId()
                if lid2 and lid2.IntegerValue > 0:
                    lvl2 = doc.GetElement(lid2)
                    if lvl2 and (lvl2.Name or "").strip():
                        return (lvl2.Name or "").strip()
        except:
            pass

    try:
        p = inst.LookupParameter("Level")
        if p and p.HasValue:
            s = (p.AsString() or "").strip()
            if s:
                return s
            vs = (p.AsValueString() or "").strip()
            if vs:
                return vs
    except:
        pass

    return "?"

def _format_opening_row(eid_int, prefix=None):
    """Row must contain 'Id <num>' for parsing. FIXED: Better error handling."""
    try:
        e = doc.GetElement(DB.ElementId(int(eid_int)))
        if not e:
            base = "Id {}".format(eid_int)
        else:
            lvl = get_level_name(e)
            pw = lookup_param_any(e, WIDTH_NAMES)
            ph = lookup_param_any(e, HEIGHT_NAMES)
            pd = lookup_param_any(e, DEPTH_NAMES)

            wmm = None
            hmm = None
            dmm = None
            
            if pw and pw.HasValue:
                try:
                    wmm = ft_to_mm(pw.AsDouble())
                except:
                    pass
            if ph and ph.HasValue:
                try:
                    hmm = ft_to_mm(ph.AsDouble())
                except:
                    pass
            if pd and pd.HasValue:
                try:
                    dmm = ft_to_mm(pd.AsDouble())
                except:
                    pass

            if wmm is not None and hmm is not None and dmm is not None:
                base = "Id {} | Level: {} | {:.0f} x {:.0f} x {:.0f} mm".format(eid_int, lvl, wmm, hmm, dmm)
            elif wmm is not None and hmm is not None:
                base = "Id {} | Level: {} | {:.0f} x {:.0f} mm".format(eid_int, lvl, wmm, hmm)
            else:
                base = "Id {} | Level: {}".format(eid_int, lvl)

        if prefix:
            return "[{}] {}".format(prefix, base)
        return base
    except Exception as ex:
        if prefix:
            return "[{}] Id {} (Error: {})".format(prefix, eid_int, str(ex))
        return "Id {} (Error: {})".format(eid_int, str(ex))

def _parse_id_from_row(row_text):
    """Parse element ID from list row."""
    try:
        s = str(row_text).strip()
        if s.startswith("[") and "]" in s:
            s = s.split("]", 1)[1].strip()
        parts = s.split()
        if len(parts) >= 2 and parts[0] == "Id":
            return int(parts[1])
    except:
        pass
    return None

def _collect_all_openings_ids_for_family(opening_family_id):
    """Collect all openings and sort by level."""
    els = collect_existing_openings(opening_family_id)
    ids = []
    for e in els:
        try:
            ids.append(e.Id.IntegerValue)
        except:
            pass

    def sort_key(i):
        try:
            e = doc.GetElement(DB.ElementId(i))
            return (get_level_name(e), i)
        except:
            return ("?", i)

    try:
        ids.sort(key=sort_key)
    except:
        pass
    return ids


class _ManagerExternalHandler(IExternalEventHandler):
    def __init__(self):
        self.action = None
        self.ids = []
        self.pick_id = None

        self.win = None
        self.lst = None
        self.cmb = None
        self.txt_counts = None

        self.family_id = None
        self.run_created = []
        self.run_updated = []

        self._ev = None

    def GetName(self):
        return "SHN Opening Manager ExternalEvent"

    def _source_index(self):
        try:
            if self.cmb is None:
                return 0
            return int(self.cmb.SelectedIndex)
        except:
            return 0

    def Execute(self, uiapp_exec):
        try:
            uidoc_exec = uiapp_exec.ActiveUIDocument
            doc_exec = uidoc_exec.Document

            action = self.action
            ids = list(self.ids) if self.ids else []
            pick_id = self.pick_id

            self.action = None
            self.ids = []
            self.pick_id = None

            if action == "refresh":
                if self.lst is None:
                    return

                src = self._source_index()

                self.lst.Items.Clear()

                if src == 0:
                    # This Run (Created + Updated)
                    for i in self.run_created:
                        self.lst.Items.Add(_format_opening_row(i, prefix="C"))
                    for i in self.run_updated:
                        self.lst.Items.Add(_format_opening_row(i, prefix="U"))

                    if self.txt_counts is not None:
                        self.txt_counts.Text = "This Run: Created {} / Updated {}".format(
                            len(self.run_created), len(self.run_updated))

                else:
                    # All in model
                    ids_all = []
                    if self.family_id is not None:
                        ids_all = _collect_all_openings_ids_for_family(self.family_id)

                    for oid in ids_all:
                        self.lst.Items.Add(_format_opening_row(oid))

                    if self.txt_counts is not None:
                        self.txt_counts.Text = "All in Model: {}".format(len(ids_all))

                # Auto-select first item
                try:
                    if self.lst.Items.Count > 0:
                        if pick_id is not None:
                            # Try to restore selection
                            for idx in range(self.lst.Items.Count):
                                if ("Id {}".format(pick_id)) in str(self.lst.Items[idx]):
                                    self.lst.SelectedIndex = idx
                                    break
                            if self.lst.SelectedIndex < 0:
                                self.lst.SelectedIndex = 0
                        else:
                            self.lst.SelectedIndex = 0
                except:
                    pass
                return

            if action == "select":
                if not ids: return
                sel = List[DB.ElementId]([DB.ElementId(int(i)) for i in ids])
                uidoc_exec.Selection.SetElementIds(sel)
                return

            if action == "show":
                if not ids: return
                sel = List[DB.ElementId]([DB.ElementId(int(i)) for i in ids])
                uidoc_exec.Selection.SetElementIds(sel)
                uidoc_exec.ShowElements(sel)
                return

            if action == "isolate":
                if not ids: return
                sel = List[DB.ElementId]([DB.ElementId(int(i)) for i in ids])
                uidoc_exec.Selection.SetElementIds(sel)
                uidoc_exec.ShowElements(sel)
                v = uidoc_exec.ActiveView
                if v:
                    v.IsolateElementsTemporary(sel)
                return

            if action == "clear_isolate":
                v = uidoc_exec.ActiveView
                try:
                    if v and v.IsTemporaryHideIsolateActive():
                        v.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)
                except:
                    pass
                return

            if action == "delete":
                if not ids: return
                t = DB.Transaction(doc_exec, "SHN Opening Manager - Delete")
                t.Start()
                try:
                    for i in ids:
                        try:
                            doc_exec.Delete(DB.ElementId(int(i)))
                        except:
                            pass
                finally:
                    t.Commit()
                self.action = "refresh"
                self._ev.Raise()
                return

        except Exception as ex:
            print("ExternalEvent Execute error: {}".format(str(ex)))

    def raise_(self, action, ids=None, pick_id=None):
        self.action = action
        self.ids = list(ids or [])
        self.pick_id = pick_id
        try:
            self._ev.Raise()
        except:
            pass


__shn_manager_window__ = None
__shn_manager_handler__ = None
__shn_manager_event__ = None


def show_opening_manager(opening_symbol, created_ids=None, updated_ids=None):
    """FIXED: Now properly initializes and displays opening list."""
    global __shn_manager_window__, __shn_manager_handler__, __shn_manager_event__

    created_ids = created_ids or []
    updated_ids = updated_ids or []

    fam_id = opening_symbol.Family.Id
    fam_name, type_name = get_symbol_names(opening_symbol)

    # Already open -> update + refresh
    if __shn_manager_window__ is not None:
        try:
            __shn_manager_window__.Activate()
            if __shn_manager_handler__:
                __shn_manager_handler__.family_id = fam_id
                __shn_manager_handler__.run_created = list(created_ids)
                __shn_manager_handler__.run_updated = list(updated_ids)
                __shn_manager_handler__.raise_("refresh")
            return
        except:
            __shn_manager_window__ = None

    win = _load_window_from_xaml(_MANAGER_XAML)

    try:
        WindowInteropHelper(win).Owner = uiapp.MainWindowHandle
    except:
        pass

    txtHeader = win.FindName("TxtHeader")
    txtInfo   = win.FindName("TxtInfo")
    cmb       = win.FindName("CmbSource")
    txtCounts = win.FindName("TxtCounts")
    lst       = win.FindName("LstOpenings")

    btnPrev    = win.FindName("BtnPrev")
    btnNext    = win.FindName("BtnNext")
    btnRefresh = win.FindName("BtnRefresh")

    btnSelect = win.FindName("BtnSelect")
    btnShow   = win.FindName("BtnShow")
    btnIso    = win.FindName("BtnIsolate")
    btnClear  = win.FindName("BtnClearIso")
    btnDelete = win.FindName("BtnDelete")
    btnClose  = win.FindName("BtnClose")

    if txtHeader:
        txtHeader.Text = "SHN Opening Manager"
    if txtInfo:
        txtInfo.Text = (
            "Family: {} | Type: {}\n"
            "Navigate openings with Prev/Next or double-click to isolate."
        ).format(fam_name, type_name)

    handler = _ManagerExternalHandler()
    ev = ExternalEvent.Create(handler)
    handler._ev = ev

    handler.win = win
    handler.lst = lst
    handler.cmb = cmb
    handler.txt_counts = txtCounts

    handler.family_id = fam_id
    handler.run_created = list(created_ids)
    handler.run_updated = list(updated_ids)

    __shn_manager_window__ = win
    __shn_manager_handler__ = handler
    __shn_manager_event__ = ev

    def current_selected_id():
        try:
            if lst and lst.SelectedItem:
                return _parse_id_from_row(lst.SelectedItem)
        except:
            pass
        return None

    def do_isolate():
        oid = current_selected_id()
        if oid is None: return
        handler.raise_("isolate", [oid])

    def do_prev():
        if not lst or lst.Items.Count == 0: return
        idx = lst.SelectedIndex
        if idx < 0: idx = 0
        lst.SelectedIndex = max(0, idx - 1)
        do_isolate()

    def do_next():
        if not lst or lst.Items.Count == 0: return
        idx = lst.SelectedIndex
        if idx < 0: idx = 0
        lst.SelectedIndex = min(lst.Items.Count - 1, idx + 1)
        do_isolate()

    def do_refresh():
        handler.raise_("refresh", pick_id=current_selected_id())

    def do_select():
        oid = current_selected_id()
        if oid is None: return
        handler.raise_("select", [oid])

    def do_show():
        oid = current_selected_id()
        if oid is None: return
        handler.raise_("show", [oid])

    def do_delete():
        oid = current_selected_id()
        if oid is None: return
        res = forms.alert("Delete opening Id {}?".format(oid), title="SHN", yes=True, no=True)
        if not res: return
        handler.raise_("delete", [oid])

    btnPrev.Click    += lambda s, a: do_prev()
    btnNext.Click    += lambda s, a: do_next()
    btnRefresh.Click += lambda s, a: do_refresh()

    btnSelect.Click  += lambda s, a: do_select()
    btnShow.Click    += lambda s, a: do_show()
    btnIso.Click     += lambda s, a: do_isolate()
    btnClear.Click   += lambda s, a: handler.raise_("clear_isolate")
    btnDelete.Click  += lambda s, a: do_delete()
    btnClose.Click   += lambda s, a: win.Close()

    if lst:
        lst.MouseDoubleClick += lambda s, a: do_isolate()

    if cmb:
        cmb.SelectionChanged += lambda s, a: handler.raise_("refresh")

    def _on_closed(sender, args):
        global __shn_manager_window__, __shn_manager_handler__, __shn_manager_event__
        __shn_manager_window__ = None
        __shn_manager_handler__ = None
        __shn_manager_event__ = None

    win.Closed += _on_closed

    win.Show()
    try:
        win.Activate()
    except:
        pass

    # FIXED: Initial refresh happens AFTER window is shown
    handler.raise_("refresh")


# -----------------------------
# Main
# -----------------------------
def main():
    forms.alert(
        "SHN Opening Browser\n\n"
        "• Walls are taken from RevitLinkInstance\n"
        "• Openings are Face-Based and always 90° to the wall\n"
        "• Supports tray and conduit packages\n"
        "• Audit: Missing / NeedResize / Empty\n\n"
        "Choose the mode.",
        title="SHN"
    )

    mode = forms.CommandSwitchWindow.show(
        ["Create/Update + Audit", "Audit only", "Open Opening Manager (existing openings)"],
        message="Run mode"
    )
    if not mode:
        return

    if mode == "Open Opening Manager (existing openings)":
        opening_symbol = pick_opening_symbol()
        show_opening_manager(opening_symbol, [], [])
        return

    opening_symbol = pick_opening_symbol()
    link_instances = pick_link_instances()

    clearance_mm   = ask_float("Clearance (mm) added to tray/conduit size", 100.0)
    cluster_gap_mm = ask_float("Cluster gap (mm) - merge nearby penetrations (e.g. 300)", 300.0)
    search_tol_mm  = ask_float("Search tolerance (mm) - match existing openings", 200.0)
    bbox_tol_mm    = ask_float("BBox tolerance (mm) - enlarge bbox when searching walls", 300.0)
    depth_extra_mm = ask_float("Depth extra (mm) - added to wall thickness for opening depth", 50.0)

    clearance_ft   = mm_to_ft(clearance_mm)
    cluster_gap_ft = mm_to_ft(cluster_gap_mm)
    search_tol_ft  = mm_to_ft(search_tol_mm)
    bbox_tol_ft    = mm_to_ft(bbox_tol_mm)
    depth_extra_ft = mm_to_ft(depth_extra_mm)

    output.print_md("## SHN Opening Browser v0.2 (FIXED)")
    fam_name, type_name = get_symbol_names(opening_symbol)
    output.print_md("- Opening symbol: **{} : {}**".format(fam_name, type_name))

    output.print_md("- Links selected: **{}**".format(len(link_instances)))
    for li in link_instances:
        output.print_md("  - {}".format(link_display(li)))

    mep = get_mep_elements()
    output.print_md("- MEP elements (tray+conduit): **{}**".format(len(mep)))

    with forms.ProgressBar(title="SHN: Caching linked walls...", cancellable=False) as pb:
        walls_cache = build_walls_cache(link_instances)
        pb.update_progress(1, 1)

    output.print_md("- Cached linked walls: **{}**".format(len(walls_cache)))
    if not walls_cache:
        forms.alert("Could not cache walls/geometry from links. Make sure links are loaded and contain walls.", exitscript=True)

    with forms.ProgressBar(title="SHN: Computing penetrations...", cancellable=True) as pb:
        pen = compute_penetrations(mep, walls_cache, clearance_ft, bbox_tol_ft)
        pb.update_progress(1, 1)

    clustered = {}
    total_raw = 0
    total_clusters = 0
    for key, rects in pen.items():
        total_raw += len(rects)
        cl = cluster_rects(rects, cluster_gap_ft)
        clustered[key] = cl
        total_clusters += len(cl)

    output.print_md("- Raw penetrations: **{}**".format(total_raw))
    output.print_md("- Clustered opening requests: **{}**".format(total_clusters))

    create_update = (mode == "Create/Update + Audit")

    created, updated, missing, need_resize, empty = place_or_update_openings(
        opening_symbol, walls_cache, clustered, depth_extra_ft, search_tol_ft, create_update
    )

    output.print_md("## Result")
    output.print_md("- Created: **{}**".format(len(created)))
    output.print_md("- Updated: **{}**".format(len(updated)))
    output.print_md("- Missing (requests): **{}**".format(len(missing)))
    output.print_md("- NeedResize (openings): **{}**".format(len(set(need_resize))))
    output.print_md("- Empty openings: **{}**".format(len(empty)))

    if created:
        output.print_md("### Created IDs")
        for i in created:
            try:
                output.print_md("- {}".format(output.linkify(DB.ElementId(int(i)))))
            except:
                output.print_md("- {}".format(i))
    if updated:
        output.print_md("### Updated IDs")
        for i in updated:
            try:
                output.print_md("- {}".format(output.linkify(DB.ElementId(int(i)))))
            except:
                output.print_md("- {}".format(i))

    forms.alert(
        "Done.\nCreated: {}\nUpdated: {}\nMissing requests: {}\nNeedResize: {}\nEmpty: {}".format(
            len(created), len(updated), len(missing), len(set(need_resize)), len(empty)
        ),
        title="SHN Opening Browser"
    )

    post = forms.CommandSwitchWindow.show(
        ["Open Opening Manager", "Finish"],
        message="Post-run actions"
    )
    if post == "Open Opening Manager":
        show_opening_manager(opening_symbol, created, updated)

if __name__ == "__main__":
    main()