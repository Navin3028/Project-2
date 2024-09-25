import win32com.client as win32

def create_Slicer(Pivottable, Field_Name, slicer_name, top_left_cell, width_cm, height_cm):
    S = Pivottable._PivotSelect(2881, 4)
    SLcache = wb.SlicerCaches.Add2(Pivottable, Field_Name, "ProdCatSlicerCache" + slicer_name)
    SL = SLcache.Slicers.Add(SlicerDestination=wb.ActiveSheet, Name="ProdCatSlicer" + slicer_name, Top=top_left_cell.Top, Left=top_left_cell.Left, Width=width_cm * 28.35, Height=height_cm * 28.35)
    return SLcache

excel_file_path = r"C:\Users\v.jeevinee\Documents\intern\pivot\Consolidate Cloud Expense December 2023.xlsx"

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Open(excel_file_path)
data_ws = wb.Worksheets['December Detail']
data_range = data_ws.UsedRange
pivot_ws = wb.Worksheets.Add()
pivot_ws.Name = 'PivotTable2'
pivot_cache = wb.PivotCaches().Create(SourceType=win32.constants.xlDatabase, SourceData=data_range)
pivot_table_range = pivot_ws.Range("C4")
pivot_table = pivot_cache.CreatePivotTable(TableDestination=pivot_table_range, TableName="MyPivotTable")
pivot_table.RowAxisLayout(win32.constants.xlTabularRow)
row_fields = ["Cloud", "Subscription", "Tag: Dept"]
value_field = "Cost"
for position, field in enumerate(row_fields, 1):
    try:
        pivot_table.PivotFields(field).Orientation = win32.constants.xlRowField
        pivot_table.PivotFields(field).Position = position
    except Exception as e:
        print(f"Error configuring row field '{field}': {e}")

try:
    pivot_table.PivotFields(value_field).Orientation = win32.constants.xlDataField
    pivot_table.PivotFields(value_field).Function = win32.constants.xlSum
    pivot_table.PivotFields(value_field).Name = f"Total {value_field}"
except Exception as e:
    print(f"Error configuring value field '{value_field}': {e}")
pivot_ws.Rows(1).RowHeight = 200
slicer_specs = [
    {"field": "Cloud", "top_left_cell": pivot_ws.Cells(1, 1), "width_cm": 5.11, "height_cm": 2.1},
    {"field": "Tag: BU", "top_left_cell": pivot_ws.Cells(1, 10), "width_cm": 5.18, "height_cm": 11.04},
    {"field": "Tag: Owner", "top_left_cell": pivot_ws.Cells(18, 1), "width_cm": 5.09, "height_cm": 15.23},
    {"field": "Tag: Environment", "top_left_cell": pivot_ws.Cells(1, 3), "width_cm": 19.94, "height_cm": 4.42},
    {"field": "Tag: Dept", "top_left_cell": pivot_ws.Cells(1, 1), "width_cm": 5.14, "height_cm": 7.06}
]

for index, slicer_spec in enumerate(slicer_specs):
    sl_cache = create_Slicer(pivot_table, slicer_spec["field"], str(index), slicer_spec["top_left_cell"], slicer_spec["width_cm"], slicer_spec["height_cm"])

    if slicer_spec["field"] == "Tag: Environment":
        sl_cache.Slicers(1).NumberOfColumns = 6
    elif slicer_spec["field"] == "Cloud":
        sl_cache.Slicers(1).NumberOfColumns = 2

wb.Save()
wb.Close()
excel.Quit()
