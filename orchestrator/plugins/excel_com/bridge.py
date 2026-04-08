"""pywin32 COM bridge — direct Excel automation on Windows.

This module runs ONLY on Windows with Excel installed.
It controls Excel.Application via COM to perform operations
that are impossible with file-based libraries:
- Pivot table creation and refresh
- Chart creation
- Formula recalculation
- VBA macro execution

Usage (Windows):
    from plugins.excel_com.bridge import ExcelBridge
    bridge = ExcelBridge()
    bridge.open("C:\\data\\report.xlsx")
    bridge.set_cell("Index", "C7", 10)
    bridge.recalculate()
    bridge.create_pivot(...)
    bridge.save()
    bridge.close()
"""

from __future__ import annotations

import json
from pathlib import Path


def _check_win32():
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


IS_WINDOWS = _check_win32()


class ExcelBridge:
    """Controls Excel via COM automation (Windows only)."""

    def __init__(self, visible: bool = False):
        if not IS_WINDOWS:
            raise RuntimeError(
                "ExcelBridge requires Windows + pywin32. "
                "Install: pip install pywin32"
            )
        import win32com.client
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = visible
        self.excel.DisplayAlerts = False
        self.wb = None
        self.path = None

    def open(self, path: str) -> dict:
        """Open a workbook."""
        self.path = str(Path(path).resolve())
        self.wb = self.excel.Workbooks.Open(self.path)
        return self._workbook_info()

    def close(self, save: bool = False) -> None:
        """Close workbook and quit Excel."""
        if self.wb:
            self.wb.Close(SaveChanges=save)
            self.wb = None
        self.excel.Quit()

    def save(self, path: str | None = None) -> str:
        """Save workbook (optionally to a new path)."""
        if path:
            self.wb.SaveAs(str(Path(path).resolve()))
        else:
            self.wb.Save()
        return json.dumps({"status": "ok", "path": path or self.path})

    # ── Cell operations ──

    def get_cell(self, sheet: str, cell: str) -> dict:
        """Read a single cell value and formula."""
        ws = self.wb.Sheets(sheet)
        c = ws.Range(cell)
        return {
            "cell": f"{sheet}!{cell}",
            "value": c.Value,
            "formula": c.Formula if c.HasFormula else None,
            "number_format": c.NumberFormat,
            "font_bold": c.Font.Bold,
        }

    def set_cell(self, sheet: str, cell: str, value) -> dict:
        """Write a value to a cell."""
        ws = self.wb.Sheets(sheet)
        ws.Range(cell).Value = value
        return {"status": "ok", "cell": f"{sheet}!{cell}", "value": value}

    def get_range(self, sheet: str, range_str: str) -> dict:
        """Read a range of cells."""
        ws = self.wb.Sheets(sheet)
        rng = ws.Range(range_str)
        values = rng.Value
        # COM returns tuple of tuples
        if values is None:
            rows = [[None]]
        elif isinstance(values, tuple):
            rows = [list(row) if isinstance(row, tuple) else [row] for row in values]
        else:
            rows = [[values]]
        return {
            "sheet": sheet,
            "range": range_str,
            "rows": rows,
            "row_count": len(rows),
        }

    def set_range(self, sheet: str, range_str: str, data: list[list]) -> dict:
        """Write a 2D array to a range."""
        ws = self.wb.Sheets(sheet)
        ws.Range(range_str).Value = data
        return {"status": "ok", "range": f"{sheet}!{range_str}"}

    # ── Formula / Calculation ──

    def recalculate(self) -> dict:
        """Force recalculation of all formulas."""
        self.excel.CalculateFull()
        return {"status": "ok", "action": "recalculated"}

    def set_formula(self, sheet: str, cell: str, formula: str) -> dict:
        """Write a formula to a cell."""
        ws = self.wb.Sheets(sheet)
        ws.Range(cell).Formula = formula
        return {"status": "ok", "cell": f"{sheet}!{cell}", "formula": formula}

    def get_formula_result(self, sheet: str, cell: str) -> dict:
        """Get calculated result of a formula cell."""
        ws = self.wb.Sheets(sheet)
        c = ws.Range(cell)
        return {
            "cell": f"{sheet}!{cell}",
            "formula": c.Formula if c.HasFormula else None,
            "calculated_value": c.Value,
        }

    # ── Pivot Table ──

    def create_pivot(
        self,
        source_sheet: str,
        source_range: str,
        dest_sheet: str,
        dest_cell: str,
        name: str,
        row_fields: list[str],
        value_fields: list[dict],  # [{"name": "Amount", "function": "sum"}]
        col_fields: list[str] | None = None,
    ) -> dict:
        """Create a pivot table using Excel COM."""
        import win32com.client
        constants = win32com.client.constants

        src_ws = self.wb.Sheets(source_sheet)
        src_rng = src_ws.Range(source_range)

        # Create pivot cache
        pc = self.wb.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=src_rng,
        )

        # Create or get destination sheet
        try:
            dst_ws = self.wb.Sheets(dest_sheet)
        except Exception:
            dst_ws = self.wb.Sheets.Add()
            dst_ws.Name = dest_sheet

        # Create pivot table
        pt = pc.CreatePivotTable(
            TableDestination=dst_ws.Range(dest_cell),
            TableName=name,
        )

        # Add row fields
        for field_name in row_fields:
            pf = pt.PivotFields(field_name)
            pf.Orientation = 1  # xlRowField

        # Add column fields
        if col_fields:
            for field_name in col_fields:
                pf = pt.PivotFields(field_name)
                pf.Orientation = 2  # xlColumnField

        # Add value fields
        FUNC_MAP = {
            "sum": -4157,      # xlSum
            "count": -4112,    # xlCount
            "average": -4106,  # xlAverage
            "max": -4136,      # xlMax
            "min": -4139,      # xlMin
        }
        for vf in value_fields:
            field_name = vf["name"]
            func = vf.get("function", "sum").lower()
            data_field = pt.AddDataField(
                pt.PivotFields(field_name),
                f"{func.capitalize()} of {field_name}",
                FUNC_MAP.get(func, -4157),
            )

        return {
            "status": "ok",
            "pivot_name": name,
            "dest": f"{dest_sheet}!{dest_cell}",
            "row_fields": row_fields,
            "value_fields": [v["name"] for v in value_fields],
        }

    def refresh_pivot(self, name: str) -> dict:
        """Refresh a pivot table by name."""
        for ws in self.wb.Sheets:
            for pt in ws.PivotTables():
                if pt.Name == name:
                    pt.RefreshTable()
                    return {"status": "ok", "pivot": name, "action": "refreshed"}
        return {"error": f"Pivot table '{name}' not found"}

    def list_pivots(self) -> list[dict]:
        """List all pivot tables in workbook."""
        pivots = []
        for ws in self.wb.Sheets:
            for pt in ws.PivotTables():
                pivots.append({
                    "name": pt.Name,
                    "sheet": ws.Name,
                    "source": pt.SourceData,
                })
        return pivots

    # ── Chart ──

    def create_chart(
        self,
        sheet: str,
        source_range: str,
        chart_type: str = "column",
        title: str = "",
        dest_sheet: str | None = None,
    ) -> dict:
        """Create a chart."""
        CHART_TYPES = {
            "column": 51,      # xlColumnClustered
            "bar": 57,         # xlBarClustered
            "line": 4,         # xlLine
            "pie": 5,          # xlPie
            "scatter": -4169,  # xlXYScatter
            "area": 1,         # xlArea
        }

        ws = self.wb.Sheets(sheet)
        src = ws.Range(source_range)

        chart_obj = ws.ChartObjects().Add(
            Left=100, Top=100, Width=500, Height=300
        )
        chart = chart_obj.Chart
        chart.SetSourceData(Source=src)
        chart.ChartType = CHART_TYPES.get(chart_type, 51)

        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = title

        return {
            "status": "ok",
            "chart_type": chart_type,
            "title": title,
            "sheet": sheet,
            "source": source_range,
        }

    # ── Sheet operations ──

    def list_sheets(self) -> list[dict]:
        """List all sheets."""
        sheets = []
        for i in range(1, self.wb.Sheets.Count + 1):
            ws = self.wb.Sheets(i)
            sheets.append({
                "name": ws.Name,
                "index": i - 1,
                "visible": ws.Visible == -1,  # xlSheetVisible
                "rows": ws.UsedRange.Rows.Count,
                "cols": ws.UsedRange.Columns.Count,
            })
        return sheets

    # ── VBA ──

    def run_macro(self, macro_name: str) -> dict:
        """Run a VBA macro."""
        try:
            result = self.excel.Run(macro_name)
            return {"status": "ok", "macro": macro_name, "result": str(result)}
        except Exception as e:
            return {"error": str(e), "macro": macro_name}

    # ── Export ──

    def export_pdf(self, output_path: str, sheets: list[str] | None = None) -> dict:
        """Export workbook or specific sheets to PDF."""
        out = str(Path(output_path).resolve())
        if sheets:
            sheet_objs = [self.wb.Sheets(s) for s in sheets]
            sheet_objs[0].Select()
            for s in sheet_objs[1:]:
                s.Select(Replace=False)
            self.excel.ActiveSheet.ExportAsFixedFormat(0, out)  # 0 = xlTypePDF
        else:
            self.wb.ExportAsFixedFormat(0, out)
        return {"status": "ok", "path": out}

    # ── Internal ──

    def _workbook_info(self) -> dict:
        sheets = self.list_sheets()
        return {
            "file": Path(self.path).name if self.path else "",
            "sheet_count": len(sheets),
            "sheets": sheets,
        }
