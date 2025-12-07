# -*- coding: utf-8 -*-
"""
ExportRooms_v4
Exports room data to CSV and HTML with English headers and Ceiling detection.
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

def get_3d_view(document):
    """Finds a suitable 3D view for Raytracing."""
    col = DB.FilteredElementCollector(document).OfClass(DB.View3D)
    for v in col:
        # We need a valid 3D view that is not a template
        if not v.IsTemplate and not v.IsAssemblyView:
            return v
    return None

def get_ceiling_info(room, view3d):
    """
    Casts a ray upwards from the room center to find ceilings.
    Returns (Has_Ceiling_Bool, Height_Value).
    """
    if not view3d or not room.Location:
        return "No 3D View", "-"

    # Point to start (Room center + 10cm up to avoid floor)
    pt_start = room.Location.Point + DB.XYZ(0, 0, 0.5) # +0.5 ft up
    pt_dir = DB.XYZ.BasisZ # Upwards

    # Setup Raytracer
    # Find ceilings (BuiltInCategory.OST_Ceilings)
    filter_ceilings = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Ceilings)
    
    try:
        # ReferenceIntersector (Filter, FindTarget, View3D)
        intersector = DB.ReferenceIntersector(filter_ceilings, DB.FindReferenceTarget.Element, view3d)
        intersector.FindReferencesInRevitLinks = True # LOOK INSIDE LINKS
        
        # Shoot the ray
        context = intersector.FindNearest(pt_start, pt_dir)
        
        if context:
            # Distance from start point to hit
            dist_ft = context.Proximity
            # Total height = dist + start_offset
            total_height_ft = dist_ft + 0.5
            total_height_m = round(total_height_ft * FT_TO_M, 2)
            return "Yes", total_height_m
        else:
            return "No", "-"
            
    except Exception:
        return "Error", "-"

def get_rooms_from_document(document, view3d):
    """Collects rooms and calculates data."""
    results = []
    try:
        collector = DB.FilteredElementCollector(document)\
                      .OfCategory(DB.BuiltInCategory.OST_Rooms)\
                      .WhereElementIsNotElementType()
                      
        for room in collector:
            if room.Area > 0 and room.Location:
                # Basic Info
                r_num = room.Number
                p_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
                r_name = p_name.AsString() if p_name else "No Name"
                r_level = room.Level.Name if room.Level else "Unknown Level"
                
                # Area
                area_sqm = round(room.Area * SQFT_TO_SQM, 2)
                
                # 3. Height from floor to slab (Room Limit Height)
                # UnboundedHeight is the explicit height of the room object
                r_height_ft = room.UnboundedHeight
                r_height_m = round(r_height_ft * FT_TO_M, 2)
                
                # 4 & 5. Ceiling Detection
                has_ceil, ceil_h = get_ceiling_info(room, view3d)
                
                results.append({
                    "Number": r_num,
                    "Name": r_name,
                    "Level": r_level,
                    "Area": area_sqm,
                    "RoomHeight": r_height_m,
                    "HasCeiling": has_ceil,
                    "CeilingHeight": ceil_h
                })
    except Exception:
        pass
        
    return results

def get_all_rooms_data():
    """Aggregates rooms from host and links."""
    all_rooms = []
    
    # Need a 3D view in the CURRENT document to run raytracing
    view3d = get_3d_view(doc)
    
    # 1. Current Model Rooms
    all_rooms.extend(get_rooms_from_document(doc, view3d))
    
    # 2. Linked Models Rooms
    links_collector = DB.FilteredElementCollector(doc)\
                        .OfCategory(DB.BuiltInCategory.OST_RvtLinks)\
                        .WhereElementIsNotElementType()
                        
    for link_instance in links_collector:
        link_doc = link_instance.GetLinkDocument()
        if link_doc:
            # We still pass the HOST view3d because Raytracer runs in Host context
            # but looks into links via FindReferencesInRevitLinks
            rooms_in_link = get_rooms_from_document(link_doc, view3d)
            all_rooms.extend(rooms_in_link)
            
    # Sort
    all_rooms.sort(key=lambda x: x["Number"])
    return all_rooms

def save_csv(data, folder, filename):
    filepath = os.path.join(folder, filename + ".csv")
    
    with io.open(filepath, mode='w', encoding='utf-8-sig') as f:
        # 1. English Headers
        header = u"Number;Name;Level;Area (m2);Room Height (m);Has Ceiling;Ceiling Height (m)\n"
        f.write(header)
        
        for row in data:
            line = u"{};{};{};{};{};{};{}\n".format(
                row["Number"],
                row["Name"],
                row["Level"],
                str(row["Area"]).replace('.', ','),
                str(row["RoomHeight"]).replace('.', ','),
                row["HasCeiling"],
                str(row["CeilingHeight"]).replace('.', ',')
            )
            f.write(line)
            
    return filepath

def save_html(data, folder, filename):
    filepath = os.path.join(folder, filename + ".html")
    
    # 1. English Headers in HTML
    html_content = u"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            tr:nth-child(even) { background-color: #f9f9f9; }
        </style>
    </head>
    <body>
        <h2>Room Schedule</h2>
        <table>
            <tr>
                <th>Number</th>
                <th>Name</th>
                <th>Level</th>
                <th>Area (m2)</th>
                <th>Room Height (m)</th>
                <th>Has Ceiling</th>
                <th>Ceiling Height (m)</th>
            </tr>
    """
    
    for row in data:
        html_content += u"<tr>"
        html_content += u"<td>{}</td>".format(row["Number"])
        html_content += u"<td>{}</td>".format(row["Name"])
        html_content += u"<td>{}</td>".format(row["Level"])
        html_content += u"<td>{}</td>".format(row["Area"])
        html_content += u"<td>{}</td>".format(row["RoomHeight"])
        html_content += u"<td>{}</td>".format(row["HasCeiling"])
        html_content += u"<td>{}</td>".format(row["CeilingHeight"])
        html_content += u"<td>{}</td>".format(row["Source"]) if "Source" in row else u""
        html_content += u"</tr>"
        
    html_content += u"</table></body></html>"
    
    with io.open(filepath, mode='w', encoding='utf-8') as html_file:
        html_file.write(html_content)
    return filepath

# --- Main Execution ---
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
# ==================================================================================    