"""
Parcel Report Generator
For use as a script tool in ArcGIS Pro
Inputs:
  - parcel_number (str): Parcel number to query
  - output_folder (str): Folder to save JPEGs and summary text
Requirements:
  - Parcel layer: 'Subject Parcel'
  - Planning districts: 'Planning District Outline' (field: PLANNAME)
  - Land use layers: under 'Map_Pro' group, by district (field: LUC_TEXT)
  - Excel file: PlanningInfo.xlsx with columns: DISTRICT PLAN, TYPE, SIZE, DESC
"""

import arcpy
import os
import pandas as pd
from arcpy import env
from arcpy import geometry as geo

# --- User Inputs ---
parcel_number = arcpy.GetParameterAsText(0)
output_folder = arcpy.GetParameterAsText(1)

# --- Constants ---
excel_path = r"T:\County Departments\Planning_Dept\District_Plans\LandUse_Master.xlsx"  # <-- Update this path
excel_sheet = 0  # Default to first sheet
parcel_layer_name = "Subject Parcel"
district_outline_name = "Planning District Outline"
map_pro_group = "Map_Pro"

# Mapping PLANNAME to land use layer name in Map_Pro
district_to_layer = {
    "Fort Lewis Mesa": "Fort Lewis Mesa",
    "Animas Valley": "Animas Valley",
    "Animas Zoning District": "Animas Zoning District",
    "Gem Village Business District": "Gem Village Business District",
    "Durango District Plan": "Durango District Plan",
    "Florida Mesa": "Florida Mesa",
    "La Posta Road": "La Posta Road",
    "Florida Road": "Florida Road",
    "Junction Creek": "Junction Creek",
    "North County": "North County",
    "Vallecito": "Vallecito",
    "West Durango": "West Durango"
}

# --- Load Excel Data ---
df = pd.read_excel(excel_path, sheet_name=excel_sheet)

# --- Reference the project ---
project = arcpy.mp.ArcGISProject("CURRENT")
map_obj = project.listMaps()[0]  # Assumes single map

# --- Get layers ---
parcel_layer = map_obj.listLayers(parcel_layer_name)[0]
district_layer = [lyr for lyr in map_obj.listLayers(district_outline_name) if lyr.isFeatureLayer][0]

# --- Step 1: Definition query and zoom ---
parcel_layer.definitionQuery = f"ParcelNumber = '{parcel_number}'"
with arcpy.da.SearchCursor(parcel_layer, ["SHAPE@", "ParcelNumber"]) as cursor:
    parcel_geom = next(cursor)[0]
parcel_extent = parcel_geom.extent

# --- Step 2: Identify planning district ---
planning_district = None
with arcpy.da.SearchCursor(district_layer, ["SHAPE@", "PLANNAME"]) as cursor:
    for row in cursor:
        if row[0].contains(parcel_geom):
            planning_district = row[1]
            break
if not planning_district:
    raise Exception("No planning district found for selected parcel.")

# --- Step 3: Identify land use designation ---
land_use_layer_name = district_to_layer.get(planning_district)
land_use_layer = map_obj.listLayers(land_use_layer_name)[0]
land_use_value = None
with arcpy.da.SearchCursor(land_use_layer, ["SHAPE@", "LUC_TEXT"]) as cursor:
    for row in cursor:
        if row[0].overlaps(parcel_geom) or row[0].contains(parcel_geom):
            land_use_value = row[1]
            break
if not land_use_value:
    raise Exception("No land use value found for selected parcel.")

# --- Step 4: Extract adjacent land uses (N/E/S/W) ---
directions = {"North": (0, 50), "East": (50, 0), "South": (0, -50), "West": (-50, 0)}
directional_land_use = {}
centroid = parcel_geom.centroid
for dir, (dx, dy) in directions.items():
    probe = geo.PointGeometry(geo.Point(centroid.centroid.X + dx, centroid.centroid.Y + dy), parcel_geom.spatialReference)
    with arcpy.da.SearchCursor(land_use_layer, ["SHAPE@", "LUC_TEXT"]) as cursor:
        for row in cursor:
            if row[0].contains(probe):
                directional_land_use[dir] = row[1]
                break
        else:
            directional_land_use[dir] = "Unknown"

# --- Step 5: Match Excel Row ---
match = df[(df["DISTRICT PLAN"].str.strip().str.lower() == planning_district.strip().lower()) &
           (df["TYPE"].str.strip().str.lower() == land_use_value.strip().lower())]

if not match.empty:
    row = match.iloc[0]
    size = row["SIZE"]
    desc = row["DESC"]
else:
    size = desc = "Not found"

# --- Step 6: Export Layouts ---
layout_exports = {
    "Layout1": {"frame": "Figure 1", "scale": 2640},
    "Layout2": {"frame": "Figure 2", "scale": None},
    "Layout3": {"frame": "Figure 3", "scale": None},
}
for layout_name, config in layout_exports.items():
    layout = project.listLayouts(layout_name)[0]
    map_frame = layout.listElements("MAPFRAME_ELEMENT", config["frame"])[0]
    map_frame.camera.setExtent(parcel_extent)
    if config["scale"]:
        map_frame.camera.scale = config["scale"]
    out_path = os.path.join(output_folder, f"{parcel_number}_{layout_name}.jpg")
    layout.exportToJPEG(out_path, resolution=300)

# --- Step 7: Build Summary Text ---
summary_text = f"""
Parcel Number: {parcel_number}
Planning District: {planning_district}
Land Use Type: {land_use_value}

From Excel:
  Size: {size}
  Description: {desc}

Adjacent Land Use:
  North: {directional_land_use['North']}
  East:  {directional_land_use['East']}
  South: {directional_land_use['South']}
  West:  {directional_land_use['West']}
"""

# Write to file
summary_path = os.path.join(output_folder, f"{parcel_number}_summary.txt")
with open(summary_path, 'w') as f:
    f.write(summary_text)

arcpy.AddMessage("Parcel report complete.")
