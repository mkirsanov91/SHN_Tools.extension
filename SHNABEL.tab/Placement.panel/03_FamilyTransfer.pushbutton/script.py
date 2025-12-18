# -*- coding: utf-8 -*-
"""
Copy level-based lighting from linked model
into host model using level-based target family.
Enhanced WPF UI with tabs.
"""

__title__ = 'Family\nTransfer'
__doc__ = 'Copy level-based lighting from linked model into host model using level-based target family.'
__author__ = 'SHNABEL digital'


from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
import System
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')
from System.Windows import Window, Application, Thickness, GridLength, GridUnitType, HorizontalAlignment, VerticalAlignment, Visibility
from System.Windows.Markup import XamlReader
from System.Windows.Controls import Grid, TextBlock, ComboBox, ColumnDefinition
from System.Windows.Media import Brushes
from System.IO import StringReader
from System.Xml import XmlReader
import sys

doc = revit.doc
uidoc = revit.uidoc


# ----------------------------------------------------------
# XAML Definition
# ----------------------------------------------------------
XAML_STRING = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SHN Family Transfer Tool" 
        Height="700" Width="900"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="5"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="5"/>
        </Style>
        <Style TargetType="ListBox">
            <Setter Property="Margin" Value="5"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="MinWidth" Value="100"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="5"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Tab Control -->
        <TabControl x:Name="MainTabControl" Grid.Row="0" Margin="10">
            
            <!-- TAB 1: SOURCE -->
            <TabItem Header="1. Source" x:Name="TabSource">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Link Selection -->
                    <GroupBox Header="Select Linked Model" Grid.Row="0" Margin="5">
                        <StackPanel>
                            <ComboBox x:Name="LinkComboBox"/>
                            <TextBlock x:Name="LinkInfoText" Text="Select a link to continue..." 
                                     FontStyle="Italic" Foreground="Gray"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <!-- Category Selection -->
                    <GroupBox Header="Select Category (Source)" Grid.Row="1" Margin="5">
                        <StackPanel>
                            <TextBox x:Name="CategorySearchBox" 
                                   Text="Search categories..."/>
                            <ListBox x:Name="CategoryListBox" 
                                   Height="150"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <!-- Family Type Selection -->
                    <GroupBox Header="Select Family Type (Source)" Grid.Row="2" Margin="5">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBox x:Name="TypeSearchBox" 
                                   Grid.Row="0"
                                   Text="Search types..."/>
                            <ListBox x:Name="TypeListBox" 
                                   Grid.Row="1"/>
                            <TextBlock x:Name="TypeInfoText" 
                                     Grid.Row="2"
                                     Text="No type selected" 
                                     FontWeight="Bold"/>
                        </Grid>
                    </GroupBox>
                    
                    <!-- Info Panel -->
                    <Border Grid.Row="3" Background="#FFF3F3F3" 
                          BorderBrush="#FFCCCCCC" BorderThickness="1" 
                          Margin="5" Padding="10">
                        <TextBlock x:Name="SourceInfoPanel" 
                                 TextWrapping="Wrap"
                                 Text="Select source link, category and family type to continue."/>
                    </Border>
                </Grid>
            </TabItem>
            
            <!-- TAB 2: LEVEL MAPPING -->
            <TabItem Header="2. Level Mapping" x:Name="TabMapping">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="5">
                        <TextBlock Text="Map each source level to a target level in your model:" 
                                 FontWeight="Bold"
                                 VerticalAlignment="Center"
                                 Margin="5"/>
                        <TextBlock x:Name="MappingStatusText" 
                                 Text="0 levels mapped" 
                                 FontWeight="Bold"
                                 Margin="20,0,0,0"
                                 Foreground="Gray"/>
                    </StackPanel>
                    
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="MappingStackPanel" Margin="5"/>
                    </ScrollViewer>
                    
                    <Border Grid.Row="2" Background="#FFF3F3F3" 
                          BorderBrush="#FFCCCCCC" BorderThickness="1" 
                          Margin="5" Padding="10">
                        <TextBlock x:Name="MappingInfoPanel" 
                                 TextWrapping="Wrap"
                                 Text="Map source levels (from link) to target levels (in current model)."/>
                    </Border>
                </Grid>
            </TabItem>
            
            <!-- TAB 3: LEVEL FILTER -->
            <TabItem Header="3. Filter Levels" x:Name="TabLevelFilter">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="5">
                        <TextBlock Text="Select which levels to place elements on:" 
                                 FontWeight="Bold"
                                 Margin="5"/>
                        <StackPanel Orientation="Horizontal" Margin="5">
                            <Button x:Name="SelectAllLevelsButton" 
                                  Content="Select All" 
                                  MinWidth="100"
                                  Margin="0,0,5,0"/>
                            <Button x:Name="DeselectAllLevelsButton" 
                                  Content="Deselect All" 
                                  MinWidth="100"/>
                            <TextBlock x:Name="LevelFilterStatusText" 
                                     Text="0 levels selected" 
                                     FontWeight="Bold"
                                     Margin="20,0,0,0"
                                     VerticalAlignment="Center"
                                     Foreground="Gray"/>
                        </StackPanel>
                    </StackPanel>
                    
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="LevelFilterStackPanel" Margin="5"/>
                    </ScrollViewer>
                    
                    <Border Grid.Row="2" Background="#FFF3F3F3" 
                          BorderBrush="#FFCCCCCC" BorderThickness="1" 
                          Margin="5" Padding="10">
                        <TextBlock x:Name="LevelFilterInfoPanel" 
                                 TextWrapping="Wrap"
                                 Text="Uncheck levels to skip placing elements on them. This allows partial placement for quick updates or corrections."/>
                    </Border>
                </Grid>
            </TabItem>
            
            <!-- TAB 4: TARGET -->
            <TabItem Header="4. Target" x:Name="TabTarget">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Target Category Selection -->
                    <GroupBox Header="Select Category (Target in Host Model)" Grid.Row="0" Margin="5">
                        <StackPanel>
                            <TextBox x:Name="TargetCategorySearchBox" 
                                   Text="Search categories..."/>
                            <ListBox x:Name="TargetCategoryListBox" 
                                   Height="150"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <!-- Target Family Type Selection -->
                    <GroupBox Header="Select Family Type (Target Level-Based)" Grid.Row="1" Margin="5">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBox x:Name="TargetTypeSearchBox" 
                                   Grid.Row="0"
                                   Text="Search types..."/>
                            <ListBox x:Name="TargetTypeListBox" 
                                   Grid.Row="1"/>
                            <TextBlock x:Name="TargetTypeInfoText" 
                                     Grid.Row="2"
                                     Text="No type selected" 
                                     FontWeight="Bold"/>
                        </Grid>
                    </GroupBox>
                    
                    <!-- Info Panel -->
                    <Border Grid.Row="2" Background="#FFF3F3F3" 
                          BorderBrush="#FFCCCCCC" BorderThickness="1" 
                          Margin="5" Padding="10">
                        <TextBlock x:Name="TargetInfoPanel" 
                                 TextWrapping="Wrap"
                                 Text="Select target family type in host model. Must be level-based."/>
                    </Border>
                </Grid>
            </TabItem>
            
            <!-- TAB 5: PREVIEW & OPTIONS -->
            <TabItem Header="5. Preview" x:Name="TabPreview">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Header="Options" Grid.Row="0" Margin="5">
                        <StackPanel>
                            <CheckBox x:Name="CopyRotationCheck" 
                                    Content="Copy rotation from source" 
                                    IsChecked="True"/>
                            <CheckBox x:Name="AdjustHeightCheck" 
                                    Content="Adjust Z-coordinate (height)" 
                                    IsChecked="True"/>
                            <CheckBox x:Name="SelectAfterCheck" 
                                    Content="Select created elements after placement" 
                                    IsChecked="True"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <GroupBox Header="Summary" Grid.Row="1" Margin="5">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <TextBlock x:Name="SummaryTextBlock" 
                                     TextWrapping="Wrap"
                                     Padding="10"
                                     FontFamily="Consolas"
                                     Text="Configure all settings in previous tabs to see summary."/>
                        </ScrollViewer>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="2" Margin="5">
                        <Button x:Name="ValidateButton" 
                              Content="Validate Settings"
                              Background="#FF4CAF50"
                              Foreground="White"
                              FontWeight="Bold"/>
                        <TextBlock x:Name="ValidationText" 
                                 TextWrapping="Wrap"
                                 Foreground="Green"
                                 FontWeight="Bold"
                                 Visibility="Collapsed"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
        </TabControl>
        
        <!-- Bottom Buttons -->
        <Border Grid.Row="1" Background="#FFF0F0F0" 
              BorderBrush="#FFCCCCCC" BorderThickness="0,1,0,0">
            <Grid Margin="10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <Button x:Name="CancelButton" 
                      Grid.Column="0"
                      Content="Cancel"/>
                
                <ProgressBar x:Name="ProgressBar" 
                           Grid.Column="1"
                           Height="25"
                           Margin="10,5"
                           Visibility="Collapsed"/>
                
                <TextBlock x:Name="StatusText"
                         Grid.Column="1"
                         Text=""
                         VerticalAlignment="Center"
                         HorizontalAlignment="Center"
                         FontStyle="Italic"
                         Foreground="Gray"/>
                
                <Button x:Name="BackButton" 
                      Grid.Column="2"
                      Content="&lt; Back"
                      IsEnabled="False"/>
                
                <Button x:Name="NextButton" 
                      Grid.Column="3"
                      Content="Next &gt;"
                      Background="#FF2196F3"
                      Foreground="White"
                      FontWeight="Bold"
                      IsEnabled="False"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"""


# ----------------------------------------------------------
# Helper Functions
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
            placement = str(fam.FamilyPlacementType)
            # Если содержит "OneLevelBased" или "WorkPlaneBased" и т.д.
            if "Level" in placement:
                return "Level-Based"
            elif "WorkPlane" in placement:
                return "Work Plane"
            else:
                return placement
        return "?"
    except:
        return "?"


# ----------------------------------------------------------
# WPF Window Class
# ----------------------------------------------------------
class FamilyTransferWindow(Window):
    def __init__(self):
        # Load XAML
        from System.Xml import XmlReader
        reader = XmlReader.Create(StringReader(XAML_STRING))
        self.window = XamlReader.Load(reader)
        
        # Data storage
        self.doc = doc
        self.uidoc = uidoc
        
        self.links_dict = {}
        self.selected_link = None
        self.link_doc = None
        self.link_transform = None
        
        self.source_categories = {}
        self.selected_source_category = None
        
        self.source_types = {}
        self.selected_source_type = None
        self.source_elements = []
        
        self.source_levels = {}
        self.level_mapping = {}  # {source_level_name: target_level_obj}
        self.level_filter_selection = {}  # {source_level_name: is_selected (bool)}
        
        self.target_levels = {}
        self.target_categories = {}
        self.target_types = {}
        self.selected_target_category = None
        self.selected_target_type = None
        
        # Get UI elements
        self.main_tab_control = self.window.FindName("MainTabControl")
        
        # Tab 1
        self.link_combo = self.window.FindName("LinkComboBox")
        self.link_info = self.window.FindName("LinkInfoText")
        self.category_search = self.window.FindName("CategorySearchBox")
        self.category_list = self.window.FindName("CategoryListBox")
        self.type_search = self.window.FindName("TypeSearchBox")
        self.type_list = self.window.FindName("TypeListBox")
        self.type_info = self.window.FindName("TypeInfoText")
        self.source_info_panel = self.window.FindName("SourceInfoPanel")
        
        # Tab 2
        self.mapping_status = self.window.FindName("MappingStatusText")
        self.mapping_stack = self.window.FindName("MappingStackPanel")
        self.mapping_info_panel = self.window.FindName("MappingInfoPanel")
        
        # Tab 3 - Level Filter
        self.select_all_levels_btn = self.window.FindName("SelectAllLevelsButton")
        self.deselect_all_levels_btn = self.window.FindName("DeselectAllLevelsButton")
        self.level_filter_status = self.window.FindName("LevelFilterStatusText")
        self.level_filter_stack = self.window.FindName("LevelFilterStackPanel")
        self.level_filter_info_panel = self.window.FindName("LevelFilterInfoPanel")
        
        # Tab 4
        self.target_cat_search = self.window.FindName("TargetCategorySearchBox")
        self.target_cat_list = self.window.FindName("TargetCategoryListBox")
        self.target_type_search = self.window.FindName("TargetTypeSearchBox")
        self.target_type_list = self.window.FindName("TargetTypeListBox")
        self.target_type_info = self.window.FindName("TargetTypeInfoText")
        self.target_info_panel = self.window.FindName("TargetInfoPanel")
        
        # Tab 5
        self.copy_rotation_check = self.window.FindName("CopyRotationCheck")
        self.adjust_height_check = self.window.FindName("AdjustHeightCheck")
        self.select_after_check = self.window.FindName("SelectAfterCheck")
        self.summary_text = self.window.FindName("SummaryTextBlock")
        self.validate_btn = self.window.FindName("ValidateButton")
        self.validation_text = self.window.FindName("ValidationText")
        
        # Buttons
        self.cancel_btn = self.window.FindName("CancelButton")
        self.back_btn = self.window.FindName("BackButton")
        self.next_btn = self.window.FindName("NextButton")
        self.progress_bar = self.window.FindName("ProgressBar")
        self.status_text = self.window.FindName("StatusText")
        
        # Wire up events using +=
        self.link_combo.SelectionChanged += self.on_link_selection_changed
        self.category_search.GotFocus += self.on_search_box_got_focus
        self.category_search.TextChanged += self.on_category_search_changed
        self.category_list.SelectionChanged += self.on_category_selection_changed
        self.type_search.GotFocus += self.on_search_box_got_focus
        self.type_search.TextChanged += self.on_type_search_changed
        self.type_list.SelectionChanged += self.on_type_selection_changed
        
        self.select_all_levels_btn.Click += self.on_select_all_levels_click
        self.deselect_all_levels_btn.Click += self.on_deselect_all_levels_click
        
        self.target_cat_search.GotFocus += self.on_search_box_got_focus
        self.target_cat_search.TextChanged += self.on_target_category_search_changed
        self.target_cat_list.SelectionChanged += self.on_target_category_selection_changed
        self.target_type_search.GotFocus += self.on_search_box_got_focus
        self.target_type_search.TextChanged += self.on_target_type_search_changed
        self.target_type_list.SelectionChanged += self.on_target_type_selection_changed
        
        self.validate_btn.Click += self.on_validate_click
        
        self.cancel_btn.Click += self.on_cancel_click
        self.back_btn.Click += self.on_back_click
        self.next_btn.Click += self.on_next_click
        
        # Initialize
        self.load_links()
        self.load_target_levels()
        self.load_target_categories()
        self.update_navigation_buttons()
        
    def on_search_box_got_focus(self, sender, e):
        """Clear placeholder text on focus"""
        if sender.Text.startswith("Search"):
            sender.Text = ""
    
    def load_links(self):
        """Load all Revit links"""
        links_collector = DB.FilteredElementCollector(self.doc).OfClass(DB.RevitLinkInstance)
        loaded_links = [l for l in links_collector if DB.RevitLinkType.IsLoaded(self.doc, l.GetTypeId())]
        
        if not loaded_links:
            self.link_info.Text = "⚠ No loaded Revit links found in the project."
            return
        
        self.links_dict = {l.Name: l for l in loaded_links}
        
        for link_name in sorted(self.links_dict.keys()):
            self.link_combo.Items.Add(link_name)
        
        self.link_info.Text = "{} link(s) available".format(len(self.links_dict))
    
    def on_link_selection_changed(self, sender, e):
        """Handle link selection"""
        if self.link_combo.SelectedItem is None:
            return
        
        link_name = str(self.link_combo.SelectedItem)
        self.selected_link = self.links_dict[link_name]
        self.link_doc = self.selected_link.GetLinkDocument()
        self.link_transform = self.selected_link.GetTotalTransform()
        
        if not self.link_doc:
            self.link_info.Text = "⚠ Selected link has no accessible document."
            self.update_navigation_buttons()
            return
        
        self.link_info.Text = "✓ Link loaded: {}".format(link_name)
        self.load_source_categories()
        self.update_navigation_buttons()
    
    def load_source_categories(self):
        """Load categories from selected link"""
        if not self.link_doc:
            return
        
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
        
        self.source_categories = {}
        for bic in cats_to_check:
            try:
                count = (DB.FilteredElementCollector(self.link_doc)
                        .OfCategory(bic)
                        .WhereElementIsNotElementType()
                        .GetElementCount())
            except:
                count = 0
            
            if count > 0:
                try:
                    c_name = DB.Category.GetCategory(self.link_doc, bic).Name
                    self.source_categories[c_name] = {"bic": bic, "count": count}
                except:
                    pass
        
        self.category_list.Items.Clear()
        for cat_name in sorted(self.source_categories.keys()):
            count = self.source_categories[cat_name]["count"]
            display = "{} ({} elements)".format(cat_name, count)
            self.category_list.Items.Add(display)
    
    def on_category_search_changed(self, sender, e):
        """Filter categories by search text"""
        search_text = self.category_search.Text.lower()
        if search_text == "search categories...":
            search_text = ""
        
        self.category_list.Items.Clear()
        for cat_name in sorted(self.source_categories.keys()):
            if search_text in cat_name.lower():
                count = self.source_categories[cat_name]["count"]
                display = "{} ({} elements)".format(cat_name, count)
                self.category_list.Items.Add(display)
    
    def on_category_selection_changed(self, sender, e):
        """Handle category selection"""
        if self.category_list.SelectedItem is None:
            return
        
        display = str(self.category_list.SelectedItem)
        cat_name = display.split(" (")[0]
        self.selected_source_category = cat_name
        
        self.load_source_types()
        self.update_navigation_buttons()
    
    def load_source_types(self):
        """Load family types from selected category"""
        if not self.link_doc or not self.selected_source_category:
            return
        
        bic = self.source_categories[self.selected_source_category]["bic"]
        
        elements = (DB.FilteredElementCollector(self.link_doc)
                   .OfCategory(bic)
                   .WhereElementIsNotElementType()
                   .ToElements())
        
        self.source_types = {}
        type_counts = {}
        
        for el in elements:
            tid = el.GetTypeId()
            if tid == DB.ElementId.InvalidElementId:
                continue
            sym = self.link_doc.GetElement(tid)
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
            
            key = "{}: {}".format(f_name, t_name)
            
            if key not in self.source_types:
                h_info = get_hosting_info(sym)
                self.source_types[key] = {"symbol": sym, "hosting": h_info, "elements": []}
                type_counts[key] = 0
            
            self.source_types[key]["elements"].append(el)
            type_counts[key] += 1
        
        self.type_list.Items.Clear()
        for key in sorted(self.source_types.keys()):
            h_info = self.source_types[key]["hosting"]
            count = type_counts[key]
            display = "[{}] {} ({} inst.)".format(h_info, key, count)
            self.type_list.Items.Add(display)
    
    def on_type_search_changed(self, sender, e):
        """Filter types by search text"""
        search_text = self.type_search.Text.lower()
        if search_text == "search types...":
            search_text = ""
        
        self.type_list.Items.Clear()
        for key in sorted(self.source_types.keys()):
            if search_text in key.lower():
                h_info = self.source_types[key]["hosting"]
                count = len(self.source_types[key]["elements"])
                display = "[{}] {} ({} inst.)".format(h_info, key, count)
                self.type_list.Items.Add(display)
    
    def on_type_selection_changed(self, sender, e):
        """Handle type selection"""
        if self.type_list.SelectedItem is None:
            return
        
        display = str(self.type_list.SelectedItem)
        # Extract key from display: "[hosting] key (count inst.)"
        parts = display.split("] ", 1)
        if len(parts) < 2:
            return
        key_with_count = parts[1]
        key = key_with_count.rsplit(" (", 1)[0]
        
        self.selected_source_type = key
        self.source_elements = self.source_types[key]["elements"]
        
        count = len(self.source_elements)
        self.type_info.Text = "✓ Selected: {} ({} instances)".format(key, count)
        
        # Load levels from source elements (also updates info panel)
        self.load_source_levels()
        
        # *** Automatically prepare mapping UI ***
        self.update_mapping_ui()
        
        # Enable Next button if everything is valid
        self.update_navigation_buttons()
    
    def load_source_levels(self):
        """Extract levels from source elements"""
        self.source_levels = {}
        
        for el in self.source_elements:
            try:
                lid = el.LevelId
                if lid != DB.ElementId.InvalidElementId:
                    lev = self.link_doc.GetElement(lid)
                    if isinstance(lev, DB.Level):
                        if lev.Name not in self.source_levels:
                            self.source_levels[lev.Name] = {"level": lev, "count": 0}
                        self.source_levels[lev.Name]["count"] += 1
            except:
                pass
        
        # Update info in source panel
        if not self.source_levels:
            self.source_info_panel.Text = "⚠ Warning: Selected elements have no valid levels!\nCannot proceed with mapping."
        else:
            info = "✓ Source Configuration Complete\n\n"
            info += "Link: {}\n".format(self.link_combo.SelectedItem)
            info += "Category: {}\n".format(self.selected_source_category)
            info += "Family Type: {}\n".format(self.selected_source_type)
            info += "Instances: {}\n".format(len(self.source_elements))
            info += "Levels found: {}".format(len(self.source_levels))
            self.source_info_panel.Text = info
    
    def load_target_levels(self):
        """Load levels from current model"""
        levels = DB.FilteredElementCollector(self.doc).OfClass(DB.Level).ToElements()
        self.target_levels = {l.Name: l for l in levels}
    

    def update_mapping_ui(self):
        """Rebuild mapping UI"""
        self.mapping_stack.Children.Clear()
        
        if not self.source_levels:
            tb = TextBlock()
            tb.Text = "No source levels available. Please complete Step 1 first."
            tb.Margin = Thickness(10)
            tb.FontStyle = System.Windows.FontStyles.Italic
            tb.Foreground = Brushes.Gray
            self.mapping_stack.Children.Add(tb)
            return
        
        for src_name in sorted(self.source_levels.keys()):
            count = self.source_levels[src_name]["count"]
            
            grid = Grid()
            grid.Margin = Thickness(5)
            
            col1 = ColumnDefinition()
            col1.Width = GridLength(1, GridUnitType.Star)
            col2 = ColumnDefinition()
            col2.Width = GridLength(30)
            col3 = ColumnDefinition()
            col3.Width = GridLength(1, GridUnitType.Star)
            
            grid.ColumnDefinitions.Add(col1)
            grid.ColumnDefinitions.Add(col2)
            grid.ColumnDefinitions.Add(col3)
            
            # Source level
            tb_src = TextBlock()
            tb_src.Text = "{} ({} inst.)".format(src_name, count)
            tb_src.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(tb_src, 0)
            grid.Children.Add(tb_src)
            
            # Arrow
            tb_arrow = TextBlock()
            tb_arrow.Text = "→"
            tb_arrow.FontSize = 16
            tb_arrow.HorizontalAlignment = HorizontalAlignment.Center
            tb_arrow.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(tb_arrow, 1)
            grid.Children.Add(tb_arrow)
            
            # Target level combo
            combo = ComboBox()
            combo.Tag = src_name  # Store source level name
            
            # Add empty option first
            combo.Items.Add("-- Select Target Level --")
            
            for tgt_name in sorted(self.target_levels.keys()):
                combo.Items.Add(tgt_name)
            
            # Set selection if mapped
            if src_name in self.level_mapping:
                mapped_name = self.level_mapping[src_name].Name
                combo.SelectedItem = mapped_name
            else:
                combo.SelectedIndex = 0  # Select placeholder
            
            combo.SelectionChanged += self.on_mapping_combo_changed
            Grid.SetColumn(combo, 2)
            grid.Children.Add(combo)
            
            self.mapping_stack.Children.Add(grid)
        
        self.update_mapping_status()
    
    def on_mapping_combo_changed(self, sender, e):
        """Handle mapping combo selection"""
        combo = sender
        src_name = str(combo.Tag)
        selected = combo.SelectedItem
        
        # Skip if nothing selected or placeholder selected
        if not selected or str(selected).startswith("--"):
            if src_name in self.level_mapping:
                del self.level_mapping[src_name]
            self.update_mapping_status()
            self.update_navigation_buttons()
            return
        
        tgt_name = str(selected)
        
        if tgt_name in self.target_levels:
            self.level_mapping[src_name] = self.target_levels[tgt_name]
        elif src_name in self.level_mapping:
            del self.level_mapping[src_name]
        
        self.update_mapping_status()
        self.update_navigation_buttons()
    
    def update_mapping_status(self):
        """Update mapping status text"""
        total = len(self.source_levels)
        mapped = len(self.level_mapping)
        self.mapping_status.Text = "{} / {} levels mapped".format(mapped, total)
        
        if mapped == total and total > 0:
            self.mapping_status.Foreground = Brushes.Green
        elif mapped > 0:
            self.mapping_status.Foreground = Brushes.Orange
        else:
            self.mapping_status.Foreground = Brushes.Gray
    
    def update_level_filter_ui(self):
        """Build level filter UI with checkboxes"""
        self.level_filter_stack.Children.Clear()
        
        if not self.level_mapping:
            tb = TextBlock()
            tb.Text = "No level mapping defined. Please complete level mapping first."
            tb.Margin = Thickness(10)
            tb.FontStyle = System.Windows.FontStyles.Italic
            tb.Foreground = Brushes.Gray
            self.level_filter_stack.Children.Add(tb)
            return
        
        from System.Windows.Controls import CheckBox
        
        # Initialize selection - all checked by default
        for src_name in self.level_mapping.keys():
            if src_name not in self.level_filter_selection:
                self.level_filter_selection[src_name] = True
        
        for src_name in sorted(self.level_mapping.keys()):
            tgt_lvl = self.level_mapping[src_name]
            count = self.source_levels[src_name]["count"]
            
            checkbox = CheckBox()
            checkbox.Content = "{} → {} ({} instances)".format(src_name, tgt_lvl.Name, count)
            checkbox.Tag = src_name
            checkbox.IsChecked = self.level_filter_selection.get(src_name, True)
            checkbox.Margin = Thickness(5)
            checkbox.FontSize = 14
            checkbox.Checked += self.on_level_filter_checkbox_changed
            checkbox.Unchecked += self.on_level_filter_checkbox_changed
            
            self.level_filter_stack.Children.Add(checkbox)
        
        self.update_level_filter_status()
    
    def on_level_filter_checkbox_changed(self, sender, e):
        """Handle checkbox state change"""
        checkbox = sender
        src_name = str(checkbox.Tag)
        self.level_filter_selection[src_name] = checkbox.IsChecked
        self.update_level_filter_status()
        self.update_navigation_buttons()
    
    def on_select_all_levels_click(self, sender, e):
        """Select all levels"""
        for child in self.level_filter_stack.Children:
            if hasattr(child, 'IsChecked'):
                child.IsChecked = True
    
    def on_deselect_all_levels_click(self, sender, e):
        """Deselect all levels"""
        for child in self.level_filter_stack.Children:
            if hasattr(child, 'IsChecked'):
                child.IsChecked = False
    
    def update_level_filter_status(self):
        """Update level filter status text"""
        if not self.level_mapping:
            self.level_filter_status.Text = "0 levels selected"
            self.level_filter_status.Foreground = Brushes.Gray
            return
        
        total = len(self.level_mapping)
        selected = sum(1 for v in self.level_filter_selection.values() if v)
        
        self.level_filter_status.Text = "{} / {} levels selected".format(selected, total)
        
        if selected == total and total > 0:
            self.level_filter_status.Foreground = Brushes.Green
        elif selected > 0:
            self.level_filter_status.Foreground = Brushes.Orange
        else:
            self.level_filter_status.Foreground = Brushes.Red
    
    def load_target_categories(self):
        """Load categories from host model"""
        all_symbols = DB.FilteredElementCollector(self.doc).OfClass(DB.FamilySymbol).ToElements()
        
        cats = {}
        for sym in all_symbols:
            try:
                if sym.Category:
                    cat_name = sym.Category.Name
                    if cat_name not in cats:
                        cats[cat_name] = {"id": sym.Category.Id, "count": 0}
                    cats[cat_name]["count"] += 1
            except:
                pass
        
        self.target_categories = cats
    
    def on_target_category_search_changed(self, sender, e):
        """Filter target categories"""
        search_text = self.target_cat_search.Text.lower()
        if search_text == "search categories...":
            search_text = ""
        
        self.target_cat_list.Items.Clear()
        for cat_name in sorted(self.target_categories.keys()):
            if search_text in cat_name.lower():
                count = self.target_categories[cat_name]["count"]
                display = "{} ({} types)".format(cat_name, count)
                self.target_cat_list.Items.Add(display)
    
    def on_target_category_selection_changed(self, sender, e):
        """Handle target category selection"""
        if self.target_cat_list.SelectedItem is None:
            return
        
        display = str(self.target_cat_list.SelectedItem)
        cat_name = display.split(" (")[0]
        self.selected_target_category = cat_name
        
        self.load_target_types()
        self.update_navigation_buttons()
    
    def load_target_types(self):
        """Load family types from selected target category"""
        if not self.selected_target_category:
            return
        
        cat_id = self.target_categories[self.selected_target_category]["id"]
        
        symbols = (DB.FilteredElementCollector(self.doc)
                  .OfCategoryId(cat_id)
                  .WhereElementIsElementType()
                  .ToElements())
        
        self.target_types = {}
        
        for sym in symbols:
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
            key = "{}: {}".format(f_name, t_name)
            display = "[{}] {}".format(h_info, key)
            
            if display in self.target_types:
                display = "{} (ID {})".format(display, sym.Id)
            
            self.target_types[display] = sym
        
        self.target_type_list.Items.Clear()
        for display in sorted(self.target_types.keys()):
            self.target_type_list.Items.Add(display)
    
    def on_target_type_search_changed(self, sender, e):
        """Filter target types"""
        search_text = self.target_type_search.Text.lower()
        if search_text == "search types...":
            search_text = ""
        
        self.target_type_list.Items.Clear()
        for display in sorted(self.target_types.keys()):
            if search_text in display.lower():
                self.target_type_list.Items.Add(display)
    
    def on_target_type_selection_changed(self, sender, e):
        """Handle target type selection"""
        if self.target_type_list.SelectedItem is None:
            return
        
        display = str(self.target_type_list.SelectedItem)
        self.selected_target_type = self.target_types[display]
        
        self.target_type_info.Text = "✓ Selected: {}".format(display)
        self.update_target_info_panel()
        self.update_navigation_buttons()
    
    def update_target_info_panel(self):
        """Update target info panel"""
        if not self.selected_target_type:
            self.target_info_panel.Text = "Select target family type in host model. Must be level-based."
            return
        
        info = "✓ Target Configuration Complete\n\n"
        info += "Category: {}\n".format(self.selected_target_category)
        info += "Family Type: {}\n".format(self.target_type_list.SelectedItem)
        
        hosting = get_hosting_info(self.selected_target_type)
        info += "Hosting: {}".format(hosting)
        
        self.target_info_panel.Text = info
    
    def update_summary(self):
        """Update summary in preview tab"""
        if not all([self.selected_source_type, self.level_mapping, self.selected_target_type]):
            self.summary_text.Text = "Configure all settings in previous tabs to see summary."
            return
        
        summary = "=" * 50 + "\n"
        summary += "PLACEMENT SUMMARY\n"
        summary += "=" * 50 + "\n\n"
        
        summary += "SOURCE:\n"
        summary += "  Link: {}\n".format(self.link_combo.SelectedItem)
        summary += "  Category: {}\n".format(self.selected_source_category)
        summary += "  Type: {}\n".format(self.selected_source_type)
        summary += "  Total instances: {}\n\n".format(len(self.source_elements))
        
        summary += "LEVEL MAPPING & SELECTION:\n"
        selected_levels = []
        skipped_levels = []
        for src_name in sorted(self.level_mapping.keys()):
            tgt_name = self.level_mapping[src_name].Name
            count = self.source_levels[src_name]["count"]
            is_selected = self.level_filter_selection.get(src_name, False)
            if is_selected:
                summary += "  ✓ {} → {} ({} inst.)\n".format(src_name, tgt_name, count)
                selected_levels.append(src_name)
            else:
                summary += "  ✗ {} → {} (SKIPPED)\n".format(src_name, tgt_name)
                skipped_levels.append(src_name)
        summary += "\n"
        
        summary += "TARGET:\n"
        summary += "  Category: {}\n".format(self.selected_target_category)
        summary += "  Type: {}\n\n".format(self.target_type_list.SelectedItem)
        
        summary += "OPTIONS:\n"
        summary += "  Copy rotation: {}\n".format("Yes" if self.copy_rotation_check.IsChecked else "No")
        summary += "  Adjust height: {}\n".format("Yes" if self.adjust_height_check.IsChecked else "No")
        summary += "  Select after: {}\n\n".format("Yes" if self.select_after_check.IsChecked else "No")
        
        # Calculate total to place (only selected levels)
        total_to_place = sum(
            self.source_levels[src_name]["count"] 
            for src_name in selected_levels
        )
        summary += "=" * 50 + "\n"
        summary += "TOTAL ELEMENTS TO PLACE: {}\n".format(total_to_place)
        if skipped_levels:
            summary += "SKIPPED LEVELS: {}\n".format(len(skipped_levels))
        summary += "=" * 50
        
        self.summary_text.Text = summary
        self.update_navigation_buttons()
    
    def on_validate_click(self, sender, e):
        """Validate settings"""
        errors = []
        
        if not self.selected_link:
            errors.append("No link selected")
        if not self.selected_source_category:
            errors.append("No source category selected")
        if not self.selected_source_type:
            errors.append("No source type selected")
        if not self.level_mapping:
            errors.append("No level mapping defined")
        if not self.selected_target_category:
            errors.append("No target category selected")
        if not self.selected_target_type:
            errors.append("No target type selected")
        
        if errors:
            
            self.validation_text.Text = "⚠ Validation failed:\n" + "\n".join("  - " + e for e in errors)
            self.validation_text.Foreground = Brushes.Red
            self.validation_text.Visibility = System.Windows.Visibility.Visible
            # Validation failed
        else:
            
            self.validation_text.Text = "✓ All settings are valid! Ready to place elements."
            self.validation_text.Foreground = Brushes.Green
            self.validation_text.Visibility = System.Windows.Visibility.Visible
            # Validation passed
    
    def validate_tab_source(self):
        """Check if source tab is complete"""
        if not self.selected_link:
            self.status_text.Text = "⚠ Please select a linked model"
            return False
        if not self.selected_source_category:
            self.status_text.Text = "⚠ Please select a source category"
            return False
        if not self.selected_source_type:
            self.status_text.Text = "⚠ Please select a source family type"
            return False
        if not self.source_levels:
            self.status_text.Text = "⚠ Selected elements have no valid levels"
            return False
        self.status_text.Text = "✓ Source configured - click Next to map levels"
        return True
    
    def validate_tab_mapping(self):
        """Check if mapping tab is complete"""
        if not self.level_mapping:
            self.status_text.Text = "⚠ Please map at least one level"
            return False
        total = len(self.source_levels)
        mapped = len(self.level_mapping)
        if mapped < total:
            self.status_text.Text = "⚠ {}/{} levels mapped - map remaining or click Next to continue".format(mapped, total)
        else:
            self.status_text.Text = "✓ All levels mapped - click Next to filter levels"
        return True
    
    def validate_tab_level_filter(self):
        """Check if level filter tab is complete"""
        if not self.level_filter_selection:
            self.status_text.Text = "⚠ No levels available for filtering"
            return False
        selected_count = sum(1 for v in self.level_filter_selection.values() if v)
        if selected_count == 0:
            self.status_text.Text = "⚠ At least one level must be selected"
            return False
        self.status_text.Text = "✓ {} level(s) selected - click Next to select target family".format(selected_count)
        return True
    
    def validate_tab_target(self):
        """Check if target tab is complete"""
        if not self.selected_target_category:
            self.status_text.Text = "⚠ Please select a target category"
            return False
        if not self.selected_target_type:
            self.status_text.Text = "⚠ Please select a target family type"
            return False
        self.status_text.Text = "✓ Target configured - click Next to review and place"
        return True
    
    def validate_tab_preview(self):
        """Check if ready to place"""
        self.status_text.Text = "✓ Ready to place - click 'Place Elements' to continue"
        return True
    
    def update_navigation_buttons(self):
        """Update back/next button states and text based on current tab"""
        current_tab = self.main_tab_control.SelectedIndex
        
        # Back button
        self.back_btn.IsEnabled = current_tab > 0
        
        # Validate current tab and enable/disable Next
        is_valid = False
        
        if current_tab == 0:  # Source
            is_valid = self.validate_tab_source()
            self.next_btn.Content = "Next: Map Levels >"
        elif current_tab == 1:  # Mapping
            is_valid = self.validate_tab_mapping()
            self.next_btn.Content = "Next: Filter Levels >"
        elif current_tab == 2:  # Level Filter
            is_valid = self.validate_tab_level_filter()
            self.next_btn.Content = "Next: Select Target >"
        elif current_tab == 3:  # Target
            is_valid = self.validate_tab_target()
            self.next_btn.Content = "Next: Preview >"
        elif current_tab == 4:  # Preview
            is_valid = self.validate_tab_preview()
            self.next_btn.Content = "Place Elements"
        
        self.next_btn.IsEnabled = is_valid
    
    def on_cancel_click(self, sender, e):
        """Cancel and close"""
        self.window.DialogResult = False
        self.window.Close()
    
    def on_back_click(self, sender, e):
        """Go to previous tab"""
        current = self.main_tab_control.SelectedIndex
        if current > 0:
            self.main_tab_control.SelectedIndex = current - 1
        self.update_navigation_buttons()
    
    def on_next_click(self, sender, e):
        """Handle Next button - navigate or place based on current tab"""
        current = self.main_tab_control.SelectedIndex
        
        if current == 4:  # Preview tab - Place elements
            self.place_families()
            return
        
        # For other tabs - navigate to next
        if current < 4:
            self.main_tab_control.SelectedIndex = current + 1
            
            # Special actions when arriving at certain tabs
            if current + 1 == 2:  # Arriving at Level Filter tab
                self.update_level_filter_ui()
            elif current + 1 == 4:  # Arriving at Preview tab
                self.update_summary()
            
            self.update_navigation_buttons()
    

    def place_families(self):
        """Main placement logic"""
        
        
        # Show progress bar
        self.progress_bar.Visibility = System.Windows.Visibility.Visible
        self.progress_bar.IsIndeterminate = True
        self.next_btn.IsEnabled = False
        
        # Activate target type
        with DB.Transaction(self.doc, "Activate target family type") as t:
            t.Start()
            if not self.selected_target_type.IsActive:
                self.selected_target_type.Activate()
                self.doc.Regenerate()
            t.Commit()
        
        count_placed = 0
        count_skipped_duplicates = 0
        count_skipped_existing = 0
        created_ids = List[DB.ElementId]()
        errors = []
        used_levels = set()
        placed_points = {}  # Track placed points to avoid duplicates: {level_name: [points]}
        
        # *** GET EXISTING ELEMENTS IN MODEL ***
        existing_elements_by_level = {}  # {level_id: [existing_elements]}
        try:
            existing_collector = (DB.FilteredElementCollector(self.doc)
                                .OfClass(DB.FamilyInstance)
                                .WhereElementIsNotElementType())
            
            for exist_el in existing_collector:
                if exist_el.Symbol.Id == self.selected_target_type.Id:
                    lvl_id = exist_el.LevelId
                    if lvl_id not in existing_elements_by_level:
                        existing_elements_by_level[lvl_id] = []
                    existing_elements_by_level[lvl_id].append(exist_el)
        except:
            pass
        
        copy_rotation = self.copy_rotation_check.IsChecked
        adjust_height = self.adjust_height_check.IsChecked
        
        with DB.Transaction(self.doc, "SHN: Copy Level-Based Families From Link") as t:
            t.Start()
            
            for el in self.source_elements:
                try:
                    fi = el if isinstance(el, DB.FamilyInstance) else None
                    if fi is None:
                        continue
                    
                    # Level
                    lid = fi.LevelId
                    if lid == DB.ElementId.InvalidElementId:
                        continue
                    lev_src = self.link_doc.GetElement(lid)
                    if not isinstance(lev_src, DB.Level):
                        continue
                    src_lvl_name = lev_src.Name
                    
                    # Check mapping
                    if src_lvl_name not in self.level_mapping:
                        continue
                    
                    # Check if level is selected in filter
                    if not self.level_filter_selection.get(src_lvl_name, False):
                        continue
                    
                    tgt_lvl = self.level_mapping[src_lvl_name]
                    tgt_lvl_name = tgt_lvl.Name
                    
                    # Geometry
                    loc = fi.Location
                    if not isinstance(loc, DB.LocationPoint):
                        continue
                    
                    pt_link = loc.Point
                    rot = loc.Rotation
                    
                    # Transform to host coordinates
                    pt_host = self.link_transform.OfPoint(pt_link)
                    
                    # *** CHECK EXISTING ELEMENTS IN MODEL ***
                    is_existing = False
                    if tgt_lvl.Id in existing_elements_by_level:
                        for exist_el in existing_elements_by_level[tgt_lvl.Id]:
                            exist_loc = exist_el.Location
                            if isinstance(exist_loc, DB.LocationPoint):
                                exist_pt = exist_loc.Point
                                dist = pt_host.DistanceTo(exist_pt)
                                if dist < 0.003:  # ~1mm tolerance
                                    is_existing = True
                                    break
                    
                    if is_existing:
                        count_skipped_existing += 1
                        continue  # Skip if element already exists in model
                    
                    # *** CHECK FOR DUPLICATES IN CURRENT BATCH ***
                    if tgt_lvl_name not in placed_points:
                        placed_points[tgt_lvl_name] = []
                    
                    # Check if element already placed in this batch (within 1mm tolerance)
                    is_duplicate = False
                    for existing_pt in placed_points[tgt_lvl_name]:
                        dist = pt_host.DistanceTo(existing_pt)
                        if dist < 0.003:  # ~1mm tolerance
                            is_duplicate = True
                            break
                    
                    if is_duplicate:
                        count_skipped_duplicates += 1
                        continue  # Skip duplicate
                    
                    # Create instance
                    new_inst = self.doc.Create.NewFamilyInstance(
                        pt_host,
                        self.selected_target_type,
                        tgt_lvl,
                        DB.Structure.StructuralType.NonStructural
                    )
                    
                    # Track this point
                    placed_points[tgt_lvl_name].append(pt_host)
                    
                    # Rotation
                    if copy_rotation:
                        try:
                            axis = DB.Line.CreateBound(
                                pt_host,
                                pt_host + DB.XYZ(0, 0, 1)
                            )
                            DB.ElementTransformUtils.RotateElement(self.doc, new_inst.Id, axis, rot)
                        except:
                            pass
                    
                    # Height adjustment
                    if adjust_height:
                        try:
                            self.doc.Regenerate()
                            loc_new = new_inst.Location
                            if isinstance(loc_new, DB.LocationPoint):
                                current_z = loc_new.Point.Z
                                diff_z = pt_host.Z - current_z
                                if abs(diff_z) > 0.001:
                                    move_vec = DB.XYZ(0, 0, diff_z)
                                    DB.ElementTransformUtils.MoveElement(self.doc, new_inst.Id, move_vec)
                        except:
                            pass
                    
                    count_placed += 1
                    created_ids.Add(new_inst.Id)
                    used_levels.add(tgt_lvl.Name)
                
                except Exception as ex:
                    errors.append(str(ex))
            
            t.Commit()
        
        # Hide progress bar
        self.progress_bar.Visibility = System.Windows.Visibility.Collapsed
        self.next_btn.IsEnabled = True
        
        # Select created elements
        if count_placed > 0 and self.select_after_check.IsChecked:
            try:
                self.uidoc.Selection.SetElementIds(created_ids)
            except:
                pass
        
        # Show result
        if count_placed > 0:
            msg = (
                "Success!\n\n"
                "Link: {}\n"
                "Source family: {}\n"
                "Target family: {}\n"
                "Created instances: {}\n"
                "Levels used: {}"
            ).format(
                self.link_combo.SelectedItem,
                self.selected_source_type,
                self.target_type_list.SelectedItem,
                count_placed,
                ", ".join(sorted(list(used_levels)))
            )
            if count_skipped_existing > 0:
                msg += "\n\nSkipped {} element(s) - already exist in model".format(count_skipped_existing)
            if count_skipped_duplicates > 0:
                msg += "\nSkipped {} duplicate(s) in source".format(count_skipped_duplicates)
            forms.alert(msg)
            self.window.DialogResult = True
            self.window.Close()
        else:
            msg = "Nothing was created."
            if count_skipped_existing > 0:
                msg += "\n\nAll {} element(s) already exist in model at same locations.".format(count_skipped_existing)
                msg += "\n\nTip: Delete existing elements first if you want to replace them."
            if count_skipped_duplicates > 0:
                msg += "\n{} duplicate(s) in source were also skipped.".format(count_skipped_duplicates)
            if errors:
                msg += "\n\nFirst error:\n{}".format(errors[0])
            forms.alert(msg, warn_icon=True)
    
    def show_dialog(self):
        """Show the window as a dialog"""
        return self.window.ShowDialog()


# ----------------------------------------------------------
# Main Execution
# ----------------------------------------------------------
if __name__ == '__main__':
    try:
        window = FamilyTransferWindow()
        window.show_dialog()
    except Exception as e:
        forms.alert("Error: {}".format(str(e)), exitscript=True)
