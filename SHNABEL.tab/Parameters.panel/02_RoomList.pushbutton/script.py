# -*- coding: utf-8 -*-
"""
ExportRooms_v9
Exports room data to CSV and HTML with English headers, Ceiling detection
(by checking ceilings' bounding boxes in the same document as the room),
and Door Count per room (from FromRoom/ToRoom across all phases).

HTML report:
- Click on column headers to sort (text / numeric).
- Filter panel with checkboxes for "Has Ceiling" and "Source".
"""

import os
import io
from pyrevit import revit, DB, forms

# --- Settings ---
BASE_PATH = r"F:\REVIT_SHN\CHECK\Rooms"
SQFT_TO_SQM = 0.09290304
FT_TO_M = 0.3048

doc = revit.doc


def get_project_info():
    """Gets clean project and model names for folder structure."""
    model_name = doc.Title
    if ".rvt" in model_name.lower():
        model_name = model_name.replace(".rvt", "").replace(".RVT", "")

    project_info = doc.ProjectInformation
    project_name = project_info.Name if project_info.Name else "Unknown_Project"

    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in invalid_chars:
        project_name = project_name.replace(char, "_")
        model_name = model_name.replace(char, "_")

    return project_name, model_name


# ---------- CEILINGS COLLECTION ----------

def collect_ceilings_bboxes(document):
    """
    Collects all ceilings in the given document and returns list of their
    bounding boxes (in document INTERNAL coordinates).
    """
    result = []
    try:
        col = (DB.FilteredElementCollector(document)
               .OfCategory(DB.BuiltInCategory.OST_Ceilings)
               .WhereElementIsNotElementType())
        for ceil in col:
            bb = ceil.get_BoundingBox(None)
            if bb:
                result.append(bb)
    except Exception as e:
        print("Error collecting ceilings in doc {}: {}".format(document.Title, e))
    return result


def get_room_center_and_floor(room):
    """
    Returns (center_point, floor_elevation_ft) for the room.
    Both values are in INTERNAL Revit coordinates (feet).

    Floor elevation берём из Min.Z bounding box'а, чтобы быть
    в той же системе координат, что и потолки.
    """
    bb = room.get_BoundingBox(None)
    if not bb:
        # fallback: use Location point
        if room.Location and hasattr(room.Location, "Point"):
            pt = room.Location.Point
            floor_z = pt.Z
            return pt, floor_z
        return None, None

    minpt = bb.Min
    maxpt = bb.Max

    center = DB.XYZ(
        (minpt.X + maxpt.X) / 2.0,
        (minpt.Y + maxpt.Y) / 2.0,
        (minpt.Z + maxpt.Z) / 2.0
    )

    # пол помещения берём по нижней отметке bounding box
    floor_z = minpt.Z

    return center, floor_z


def find_ceiling_above_room(room, ceilings_bboxes):
    """
    Для заданной комнаты и списка bbox потолков (в том же документе)
    находит ближайший потолок над центром комнаты по Z.

    Возвращает: (HasCeilingStr, Height_m_or_dash)
    Height = расстояние от пола комнаты (bb.Min.Z) до низа потолка.
    """
    center, floor_z = get_room_center_and_floor(room)
    if center is None:
        return "No", "-"

    tol_xy = 0.1  # small tolerance in ft
    closest_ceil_z = None

    for bb in ceilings_bboxes:
        cmin = bb.Min
        cmax = bb.Max

        # check XY: center inside ceiling footprint (with small tolerance)
        if not (cmin.X - tol_xy <= center.X <= cmax.X + tol_xy and
                cmin.Y - tol_xy <= center.Y <= cmax.Y + tol_xy):
            continue

        # ceiling bottom elevation (internal coords, feet)
        ceil_z = cmin.Z

        # must be above floor
        if ceil_z <= floor_z + 0.1:
            continue

        if closest_ceil_z is None or ceil_z < closest_ceil_z:
            closest_ceil_z = ceil_z

    if closest_ceil_z is None:
        return "No", "-"

    height_m = round((closest_ceil_z - floor_z) * FT_TO_M, 2)
    return "Yes", height_m


# ---------- DOORS -> ROOM COUNTS ----------

def build_door_room_counts(document):
    """
    Строит словарь {roomId (int) : number_of_doors} для данного документа.

    Использует FromRoom/ToRoom по всем фазам документа:
    идём от "последней" фазы к первой и берём первую фазу, где
    FromRoom или ToRoom не None. Дверь считается для обеих комнат.
    """
    counts = {}

    # список фаз документа
    try:
        phases = [ph for ph in document.Phases]
    except Exception as e:
        print("Error getting phases for doc {}: {}".format(document.Title, e))
        return counts

    if not phases:
        return counts

    try:
        doors = (DB.FilteredElementCollector(document)
                 .OfCategory(DB.BuiltInCategory.OST_Doors)
                 .WhereElementIsNotElementType())
    except Exception as e:
        print("Error collecting doors in doc {}: {}".format(document.Title, e))
        return counts

    for door in doors:
        try:
            rooms_for_door = []
            fr = None
            tr = None

            # перебираем фазы с конца (новые → старые)
            for ph in reversed(phases):
                try:
                    fr = door.FromRoom[ph]
                except:
                    fr = None
                try:
                    tr = door.ToRoom[ph]
                except:
                    tr = None

                if fr or tr:
                    # нашли фазу, где дверь привязана к комнатам
                    break

            if fr:
                rooms_for_door.append(fr)
            if tr and tr != fr:
                rooms_for_door.append(tr)

            for rm in rooms_for_door:
                rid = rm.Id.IntegerValue
                counts[rid] = counts.get(rid, 0) + 1

        except Exception as e_door:
            print("Error processing door {} in {}: {}".format(door.Id, document.Title, e_door))

    return counts


# ---------- ROOMS COLLECTION ----------

def get_rooms_from_document(document, ceilings_bboxes, door_counts, source_label):
    """
    Collects rooms from 'document' and calculates data,
    using 'ceilings_bboxes' and 'door_counts' (по room.Id) из того же документа.
    """
    results = []
    try:
        collector = (
            DB.FilteredElementCollector(document)
            .OfCategory(DB.BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
        )

        for room in collector:
            try:
                if room.Area <= 0 or not room.Location:
                    continue

                # Basic Info
                r_num = room.Number
                p_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
                r_name = p_name.AsString() if p_name else "No Name"
                r_level = room.Level.Name if room.Level else "Unknown Level"

                # Area (sqft -> sqm)
                area_sqm = round(room.Area * SQFT_TO_SQM, 2)

                # Room height (UnboundedHeight, ft -> m)
                r_height_ft = room.UnboundedHeight
                r_height_m = round(r_height_ft * FT_TO_M, 2)

                # Ceiling detection within this document
                has_ceil, ceil_h = find_ceiling_above_room(room, ceilings_bboxes)

                # Door count for this room
                door_count = door_counts.get(room.Id.IntegerValue, 0)

                results.append({
                    "Number": r_num,
                    "Name": r_name,
                    "Level": r_level,
                    "Area": area_sqm,
                    "RoomHeight": r_height_m,
                    "DoorCount": door_count,
                    "HasCeiling": has_ceil,
                    "CeilingHeight": ceil_h,
                    "Source": source_label
                })
            except Exception as e_room:
                print("Error processing room {} in {}: {}".format(room.Id, source_label, e_room))

    except Exception as e:
        print("Error collecting rooms in {}: {}".format(source_label, e))

    return results


def get_all_rooms_data():
    """Aggregates rooms from host and links."""
    all_rooms = []

    # 1. Host document
    host_ceilings = collect_ceilings_bboxes(doc)
    host_door_counts = build_door_room_counts(doc)
    all_rooms.extend(get_rooms_from_document(doc, host_ceilings, host_door_counts, "Host model"))

    # 2. Linked documents (rooms + ceilings + doors внутри линка)
    links_collector = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.RevitLinkInstance)
        .WhereElementIsNotElementType()
    )

    for link_instance in links_collector:
        link_doc = link_instance.GetLinkDocument()
        if not link_doc:
            continue

        source_label = "Link: {}".format(link_doc.Title)
        link_ceilings = collect_ceilings_bboxes(link_doc)
        link_door_counts = build_door_room_counts(link_doc)
        rooms_in_link = get_rooms_from_document(
            link_doc,
            link_ceilings,
            link_door_counts,
            source_label
        )
        all_rooms.extend(rooms_in_link)

    # Sort by room number and then by source
    all_rooms.sort(key=lambda x: (x["Number"], x["Source"]))
    return all_rooms


# ---------- SAVE FUNCTIONS ----------

def save_csv(data, folder, filename):
    filepath = os.path.join(folder, filename + ".csv")

    with io.open(filepath, mode='w', encoding='utf-8-sig') as f:
        header = (
            u"Number;Name;Level;Area (m2);Room Height (m);Door Count;"
            u"Has Ceiling;Ceiling Height (m);Source\n"
        )
        f.write(header)

        for row in data:
            line = u"{};{};{};{};{};{};{};{};{}\n".format(
                row["Number"],
                row["Name"],
                row["Level"],
                str(row["Area"]).replace('.', ','),          # decimal comma
                str(row["RoomHeight"]).replace('.', ','),
                row["DoorCount"],
                row["HasCeiling"],
                str(row["CeilingHeight"]).replace('.', ','),
                row["Source"].replace(";", ",")
            )
            f.write(line)

    return filepath


def save_html(data, folder, filename):
    filepath = os.path.join(folder, filename + ".html")

    html_content = u"""
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: Arial, sans-serif; }
        #filters {
            margin-bottom: 10px;
            padding: 6px;
            border: 1px solid #ddd;
            background-color: #fafafa;
        }
        .filter-group {
            display: inline-block;
            margin-right: 20px;
            margin-bottom: 4px;
        }
        .filter-group span {
            font-weight: bold;
            margin-right: 4px;
        }
        table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th {
            background-color: #f2f2f2;
            cursor: pointer;
        }
        tr:nth-child(even) { background-color: #f9f9f9; }
    </style>
</head>
<body>
    <h2>Room Schedule</h2>
    <div id="filters"></div>
    <table id="roomTable">
        <thead>
            <tr>
                <th>Number</th>
                <th>Name</th>
                <th>Level</th>
                <th>Area (m2)</th>
                <th>Room Height (m)</th>
                <th>Door Count</th>
                <th>Has Ceiling</th>
                <th>Ceiling Height (m)</th>
                <th>Source</th>
            </tr>
        </thead>
        <tbody>
"""

    for row in data:
        html_content += u"<tr>"
        html_content += u"<td>{}</td>".format(row["Number"])
        html_content += u"<td>{}</td>".format(row["Name"])
        html_content += u"<td>{}</td>".format(row["Level"])
        html_content += u"<td>{}</td>".format(row["Area"])
        html_content += u"<td>{}</td>".format(row["RoomHeight"])
        html_content += u"<td>{}</td>".format(row["DoorCount"])
        html_content += u"<td>{}</td>".format(row["HasCeiling"])
        html_content += u"<td>{}</td>".format(row["CeilingHeight"])
        html_content += u"<td>{}</td>".format(row["Source"])
        html_content += u"</tr>"

    # закрываем tbody и table, добавляем JS
    html_content += u"""
        </tbody>
    </table>

    <script>
    document.addEventListener('DOMContentLoaded', function() {
        var table = document.getElementById('roomTable');
        var tbody = table.tBodies[0];
        var rows = Array.prototype.slice.call(tbody.rows);

        // ---------- SORTING ----------
        function sortByColumn(colIndex, isNumeric) {
            var sorted = rows.slice().sort(function(a, b) {
                var aText = a.cells[colIndex].textContent.trim();
                var bText = b.cells[colIndex].textContent.trim();
                if (isNumeric) {
                    var aNum = parseFloat(aText.replace(',', '.')) || 0;
                    var bNum = parseFloat(bText.replace(',', '.')) || 0;
                    return aNum - bNum;
                } else {
                    return aText.localeCompare(bText);
                }
            });

            // toggle direction
            var currentCol = table.getAttribute('data-sort-col');
            var currentDir = table.getAttribute('data-sort-dir');
            if (currentCol == colIndex.toString() && currentDir == 'asc') {
                sorted.reverse();
                table.setAttribute('data-sort-dir', 'desc');
            } else {
                table.setAttribute('data-sort-col', colIndex);
                table.setAttribute('data-sort-dir', 'asc');
            }

            // re-append rows
            sorted.forEach(function(row) {
                tbody.appendChild(row);
            });

            // обновляем массив rows (новый порядок)
            rows = Array.prototype.slice.call(tbody.rows);
        }

        var headers = table.tHead.rows[0].cells;
        for (var i = 0; i < headers.length; i++) {
            (function(index) {
                headers[index].addEventListener('click', function() {
                    // numeric columns: Area, Room Height, Door Count, Ceiling Height
                    var numericCols = [3, 4, 5, 7];
                    var isNumeric = numericCols.indexOf(index) !== -1;
                    sortByColumn(index, isNumeric);
                });
            })(i);
        }

        // ---------- FILTERS ----------
        var filterContainer = document.getElementById('filters');
        // Has Ceiling (col 6), Source (col 8)
        var filterColumns = { 'Has Ceiling': 6, 'Source': 8 };
        var activeFilters = {};

        function buildFilters() {
            for (var label in filterColumns) {
                if (!filterColumns.hasOwnProperty(label))
                    continue;
                var colIndex = filterColumns[label];
                var valuesSet = {};

                rows.forEach(function(row) {
                    var val = row.cells[colIndex].textContent.trim();
                    if (val === '')
                        return;
                    valuesSet[val] = true;
                });

                var groupDiv = document.createElement('div');
                groupDiv.className = 'filter-group';
                var titleSpan = document.createElement('span');
                titleSpan.textContent = label + ': ';
                groupDiv.appendChild(titleSpan);

                for (var val in valuesSet) {
                    if (!valuesSet.hasOwnProperty(val))
                        continue;

                    var id = 'filter_' +
                             label.replace(/\\s+/g,'_') + '_' +
                             val.replace(/[^a-zA-Z0-9]+/g,'_');

                    var cb = document.createElement('input');
                    cb.type = 'checkbox';
                    cb.checked = true;
                    cb.value = val;
                    cb.setAttribute('data-col', colIndex);
                    cb.id = id;
                    cb.addEventListener('change', applyFilters);

                    var lb = document.createElement('label');
                    lb.htmlFor = id;
                    lb.textContent = val;

                    groupDiv.appendChild(cb);
                    groupDiv.appendChild(lb);
                }

                filterContainer.appendChild(groupDiv);
            }
        }

        function applyFilters() {
            activeFilters = {};
            var checkboxes = filterContainer.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(function(cb) {
                var col = cb.getAttribute('data-col');
                if (!(col in activeFilters))
                    activeFilters[col] = [];
                if (cb.checked)
                    activeFilters[col].push(cb.value);
            });

            rows.forEach(function(row) {
                var visible = true;
                for (var col in activeFilters) {
                    if (!activeFilters.hasOwnProperty(col))
                        continue;
                    var allowedValues = activeFilters[col];
                    if (allowedValues.length === 0) {
                        visible = false;
                        break;
                    }
                    var cellText = row.cells[parseInt(col)].textContent.trim();
                    if (allowedValues.indexOf(cellText) === -1) {
                        visible = false;
                        break;
                    }
                }
                row.style.display = visible ? '' : 'none';
            });
        }

        buildFilters();
    });
    </script>
</body>
</html>
"""

    with io.open(filepath, mode='w', encoding='utf-8') as html_file:
        html_file.write(html_content)
    return filepath


# ---------- MAIN ----------

try:
    project_name, model_name = get_project_info()
    output_dir = os.path.join(BASE_PATH, project_name, model_name)

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    data = get_all_rooms_data()

    if data:
        save_csv(data, output_dir, "Room_Schedule")
        save_html(data, output_dir, "Room_Schedule")

        msg = "Done!\nFolder: {}\nRooms found: {}".format(output_dir, len(data))
        forms.alert(msg, title="Success")
        os.startfile(output_dir)
    else:
        forms.alert("No placed rooms found.", title="Warning")

except Exception as e:
    forms.alert("Error:\n{}".format(str(e)), title="Error")
#===========================================