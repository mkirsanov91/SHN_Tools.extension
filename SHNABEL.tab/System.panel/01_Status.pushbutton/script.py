# -*- coding: utf-8 -*-
"""
Check Auto-Export status and open report folder.
"""
__title__ = 'System\nStatus'
__author__ = 'SHNABEL Dept'

from pyrevit import forms, script, revit
import os
import re

# =========================================================================
# 1. PATH SETTINGS
# =========================================================================

# Path to SHN_Tools.extension
BUTTON_DIR = os.path.dirname(__file__)       
PANEL_DIR = os.path.dirname(BUTTON_DIR)      
TAB_DIR = os.path.dirname(PANEL_DIR)         
EXTENSION_DIR = os.path.dirname(TAB_DIR)     

# Path to Hook file
HOOK_FILE = os.path.join(EXTENSION_DIR, 'hooks', 'doc-synced.py')

# Path to Server Reports
SERVER_ROOT_PATH = r"F:\REVIT_SHN\CHECK\Parameters_BOQ"

# =========================================================================
# 2. FUNCTIONS
# =========================================================================

def clean_filename(text):
    return re.sub(r'[\\/*?:"<>|]', '_', text).strip()

def get_current_project_folder():
    """Calculates report path for current project"""
    doc = revit.doc
    if not doc: return None
    
    try:
        p_info = doc.ProjectInformation
        p_name = p_info.Name if p_info.Name else "Unassigned_Project"
        
        m_title = doc.Title
        if m_title.lower().endswith('.rvt'): m_title = m_title[:-4]
        
        safe_p = clean_filename(p_name)
        safe_m = clean_filename(m_title)
        
        return os.path.join(SERVER_ROOT_PATH, safe_p, safe_m)
    except:
        return SERVER_ROOT_PATH

# =========================================================================
# 3. BUTTON UI
# =========================================================================

# Check if hook file exists
if os.path.exists(HOOK_FILE):
    status_msg = "ACTIVE"
    # Using text instead of emoji to avoid encoding issues
    status_symbol = "[ON]" 
else:
    status_msg = "DISABLED (File not found)"
    status_symbol = "[OFF]"

# Get current project folder
project_folder = get_current_project_folder()

# Main Menu Alert
res = forms.alert(
    "Auto-Export System: {} {}\n\n".format(status_msg, status_symbol) +
    "Script location:\n{}".format(HOOK_FILE),
    title="SHNABEL Control Panel",
    options=["Open Report Folder", "Open Script Folder", "Cancel"],
    footer="Ver 1.0"
)

# Handle actions
if res == "Open Report Folder":
    if project_folder and os.path.exists(project_folder):
        os.startfile(project_folder)
    else:
        # If specific folder missing, try opening root
        if os.path.exists(SERVER_ROOT_PATH):
            os.startfile(SERVER_ROOT_PATH)
            forms.alert("Folder for this model does not exist yet (Sync required).\nOpened root folder instead.")
        else:
            forms.alert("Server path not found!")

elif res == "Open Script Folder":
    os.startfile(EXTENSION_DIR)