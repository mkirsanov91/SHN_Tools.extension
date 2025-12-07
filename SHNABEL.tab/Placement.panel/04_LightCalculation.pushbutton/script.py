# -*- coding: utf-8 -*-

from Autodesk.Revit.UI.Selection import ObjectType
import math


from pyrevit import script
output = script.get_output()
output.show() # <--- ЭТА КОМАНДА ОТКРЫВАЕТ ОКНО КОНСОЛИ
print("Скрипт запущен...")


from System.Collections.Generic import List

from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc

# --- КОНФИГУРАЦИЯ ---
WORK_PLANE_HEIGHT_MM = 800  # Высота рабочей плоскости (мм)
GRID_SIZE_MM = 500          # Шаг сетки (мм)
ANALYSIS_NAME = "Анализ освещенности (Lambert)"

# Конвертация единиц
FT_TO_M = 0.3048
MM_TO_FT = 1 / 304.8

def setup_style(view, min_val=0, max_val=500):
    """
    Создает или обновляет стиль отображения (Цвета и Легенду).
    """
    style_name = "Heatmap_Lux_Style"
    
    # Пытаемся найти существующий стиль
    collector = DB.FilteredElementCollector(doc).OfClass(DB.Analysis.AnalysisDisplayStyle)
    style = None
    for s in collector:
        if s.Name == style_name:
            style = s
            break
            
    # Настройка цветов (Синий -> Зеленый -> Желтый -> Красный)
    # Создаем градиент
    color_settings = DB.Analysis.AnalysisDisplayColorSettings()
    color_settings.MinColor = DB.Color(0, 0, 255)   # Синий (0 Lux)
    color_settings.MaxColor = DB.Color(255, 0, 0)   # Красный (Max Lux)
    
    # Настройка легенды
    legend_settings = DB.Analysis.AnalysisDisplayLegendSettings()
    legend_settings.ShowLegend = True
    legend_settings.NumberOfSteps = 10 # Количество делений
    legend_settings.ShowDataDescription = False
    
    # Если стиля нет - создаем, если есть - обновляем
    t_style = DB.Transaction(doc, "Настройка стиля")
    t_style.Start()
    try:
        if not style:
            style = DB.Analysis.AnalysisDisplayStyle.CreateAnalysisDisplayStyle(
                doc, style_name, color_settings, legend_settings, legend_settings
            )
        else:
            style.SetColorSettings(color_settings)
            style.SetLegendSettings(legend_settings)
        
        # Применяем стиль к виду
        view.AnalysisDisplayStyleId = style.Id
    except Exception as e:
        print("Ошибка стиля: {}".format(e))
    t_style.Commit()
    
    return style

# --- 1. ВЫБОР И ВВОД ДАННЫХ ---
print("Пожалуйста, выберите помещение в Revit...")

# Убираем try/except, чтобы видеть ошибку, если она есть
sel_ref = uidoc.Selection.PickObject(ObjectType.Element, "Выберите помещение")
room = doc.GetElement(sel_ref)

# Проверка, что это комната
if not isinstance(room, DB.Architecture.Room):
    forms.alert("Вы тыкнули не в Room (Помещение)!", exitscript=True)

print("Помещение выбрано: {}".format(room.Name))

# Ввод люменов
res_lumens = forms.ask_for_string(
    default='3000',
    prompt='Введите световой поток светильника (Люмен):',
    title='Параметры Ламберта'
)

if not res_lumens or not res_lumens.isdigit():
    script.exit()
    
LUMENS_PER_FIXTURE = float(res_lumens)

# --- 2. ПОДГОТОВКА ГЕОМЕТРИИ ---

# Находим светильники в комнате (грубый поиск по BBox)
bbox = room.get_BoundingBox(None)
outline = DB.Outline(bbox.Min, bbox.Max)
bb_filter = DB.BoundingBoxIntersectsFilter(outline)

collector = DB.FilteredElementCollector(doc)\
    .OfCategory(DB.BuiltInCategory.OST_LightingFixtures)\
    .WhereElementIsNotElementType()\
    .WherePasses(bb_filter)

room_lights_pos = []
count_lights = 0
for light in collector:
    # Дополнительная проверка: точка светильника внутри комнаты?
    pt = light.Location.Point
    if room.IsPointInRoom(pt):
        room_lights_pos.append(pt)
        count_lights += 1

if count_lights == 0:
    forms.alert("В этом помещении не найдены светильники (проверьте уровень размещения).", exitscript=True)

# --- 3. ГЕНЕРАЦИЯ ТОЧЕК И РАСЧЕТ ---

points = []
values = []

# Высота рабочей плоскости в футах
z_plane = bbox.Min.Z + (WORK_PLANE_HEIGHT_MM * MM_TO_FT)
step_ft = GRID_SIZE_MM * MM_TO_FT

x_min, x_max = bbox.Min.X, bbox.Max.X
y_min, y_max = bbox.Min.Y, bbox.Max.Y

# Интенсивность (Candela) для изотропного источника
# I = Flux / 4Pi
intensity_candela = LUMENS_PER_FIXTURE / (4 * math.pi)

print("Расчет точек... Светильников: {}".format(count_lights))

x = x_min
while x < x_max:
    y = y_min
    while y < y_max:
        pt_calc = DB.XYZ(x, y, z_plane)
        
        if room.IsPointInRoom(pt_calc):
            total_lux = 0.0
            
            for light_pos in room_lights_pos:
                # Вектор от точки к свету
                vec = light_pos - pt_calc
                dist_ft = vec.GetLength()
                dist_m = dist_ft * FT_TO_M
                
                # Защита от деления на ноль (если точка прямо в лампе)
                if dist_m < 0.1: 
                    dist_m = 0.1
                
                # Косинус угла падения (alpha)
                # Нормаль пола (0,0,1). Угол между вектором ВВЕРХ (к свету) и нормалью.
                # cos(alpha) = Z_component / Length
                cos_alpha = vec.Z / dist_ft 
                if cos_alpha < 0: cos_alpha = 0 # Свет снизу не считаем
                
                # E = (I * cos(a)) / R^2
                lux = (intensity_candela * cos_alpha) / (dist_m ** 2)
                total_lux += lux
            
            # Сохраняем результат
            points.append(pt_calc)
            
            # Revit AVF принимает значение как ValueAtPoint
            val = DB.Analysis.ValueAtPoint([total_lux])
            values.append(val)
            
    y += step_ft
x += step_ft

if not points:
    forms.alert("Точки расчета не попали внутрь комнаты. Проверьте границы.", exitscript=True)

# --- 4. ВИЗУАЛИЗАЦИЯ (AVF) ---

t = DB.Transaction(doc, "Расчет освещенности")
t.Start()

# Получаем менеджер для текущего вида
sfm = DB.Analysis.SpatialFieldManager.GetSpatialFieldManager(doc.ActiveView)
if not sfm:
    sfm = DB.Analysis.SpatialFieldManager.CreateSpatialFieldManager(doc.ActiveView, 1)

sfm.Clear() # Очистить старые результаты

# Регистрируем схему результата (метаданные)
schema_name = "Lux Calculation"
registered_results = sfm.GetRegisteredResults()
schema_idx = 0
found = False

for i in registered_results:
    s = sfm.GetResultSchema(i)
    if s.Name == schema_name:
        schema_idx = i
        found = True
        break

if not found:
    schema = DB.Analysis.AnalysisResultSchema(schema_name, "Illuminance Analysis")
    schema_idx = sfm.RegisterResult(schema)

# Создаем примитив (облако точек)
idx_primitive = sfm.AddSpatialFieldPrimitive()

# Подготовка данных для API (требует IList)
p_list = List[DB.XYZ](points)
v_list = List[DB.Analysis.ValueAtPoint](values)

field_points = DB.Analysis.FieldDomainPointsByXYZ(p_list)
field_values = DB.Analysis.FieldValues(v_list)

# Записываем данные
sfm.UpdateSpatialFieldPrimitive(idx_primitive, field_points, field_values, schema_idx)

# Применяем красивый стиль
setup_style(doc.ActiveView, min_val=0, max_val=500)

t.Commit()

print("Готово! Рассчитано точек: {}. Макс. значение: {:.1f} Lux".format(len(points), max([v.GetValues()[0] for v in values])))