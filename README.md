# Revit Parameter Extractor — Power BI

Extracts **all instance + type parameters** from every Revit element and writes a typed **Parquet file** ready to be loaded directly in Power BI for KPI dashboards and completion-status tracking.

## Requirements

```bash
pip install pyarrow pandas
```

Run inside **pyRevit (CPython 3.x engine)** or a **Dynamo Python node**.

---

## Output formats

Set `CONFIG["output_format"]` to `"LONG"` or `"WIDE"`.

### LONG — one row per (element × parameter) ← recommended for Power BI

Best for slicers, completion KPIs, and parameter-level filtering.

| ElementId | Category | FamilyName   | TypeName                    | LevelName | WorksetName  | Phase            | ParameterName      | GroupName    | IsInstance | StorageType | DisplayValue  | RawValue      | Unit | CompletionStatus |
|-----------|----------|--------------|-----------------------------|-----------|--------------|------------------|--------------------|--------------|------------|-------------|---------------|---------------|------|-----------------|
| 123456    | Walls    | Basic Wall   | Exterior - 300mm Concrete   | Level 1   | Architecture | New Construction | Comments           | Identity Data | True       | String      | Load bearing  | Load bearing  |      | Filled          |
| 123456    | Walls    | Basic Wall   | Exterior - 300mm Concrete   | Level 1   | Architecture | New Construction | Mark               | Identity Data | True       | String      | W-01          | W-01          |      | Filled          |
| 123456    | Walls    | Basic Wall   | Exterior - 300mm Concrete   | Level 1   | Architecture | New Construction | Length             | Dimensions    | True       | Double      | 5.4           | 5.4           | m    | Filled          |
| 123456    | Walls    | Basic Wall   | Exterior - 300mm Concrete   | Level 1   | Architecture | New Construction | Area               | Dimensions    | True       | Double      | 18.9          | 18.9          | m2   | Filled          |
| 123456    | Walls    | Basic Wall   | Exterior - 300mm Concrete   | Level 1   | Architecture | New Construction | Fire Rating        | Identity Data | False      | String      |               |               |      | **Missing**     |
| 123456    | Walls    | Basic Wall   | Exterior - 300mm Concrete   | Level 1   | Architecture | New Construction | Structural Usage   | Structural    | True       | Integer     | Bearing       | 1             |      | Filled          |
| 234567    | Doors    | Single-Flush | 0915 x 2134mm               | Level 1   | Architecture | New Construction | Mark               | Identity Data | True       | String      | D-01          | D-01          |      | Filled          |
| 234567    | Doors    | Single-Flush | 0915 x 2134mm               | Level 1   | Architecture | New Construction | Comments           | Identity Data | True       | String      |               |               |      | **Missing**     |
| 234567    | Doors    | Single-Flush | 0915 x 2134mm               | Level 1   | Architecture | New Construction | Fire Rating        | Identity Data | False      | String      | 60 min        | 60 min        |      | Filled          |
| 234567    | Doors    | Single-Flush | 0915 x 2134mm               | Level 1   | Architecture | New Construction | Width              | Dimensions    | False      | Double      | 0.915         | 0.915         | m    | Filled          |

> Sample file: [`sample_long.parquet`](sample_long.parquet)

---

### WIDE — one row per element, one column per parameter

Better for flat schedule-style tables.

| ElementId | Category | FamilyName   | TypeName                    | LevelName | WorksetName  | Phase            | Mark  | Comments     | Length | Area | Fire Rating | Structural Usage | Width |
|-----------|----------|--------------|-----------------------------|-----------|--------------|------------------|-------|--------------|--------|------|-------------|-----------------|-------|
| 123456    | Walls    | Basic Wall   | Exterior - 300mm Concrete   | Level 1   | Architecture | New Construction | W-01  | Load bearing | 5.4    | 18.9 |             | Bearing          |       |
| 234567    | Doors    | Single-Flush | 0915 x 2134mm               | Level 1   | Architecture | New Construction | D-01  |              |        |      | 60 min      |                  | 0.915 |

> Sample file: [`sample_wide.parquet`](sample_wide.parquet)

---

## Power BI KPI ideas (LONG format)

| Measure | DAX |
|---------|-----|
| Completion % | `DIVIDE(COUNTROWS(FILTER(Table, Table[CompletionStatus]="Filled")), COUNTROWS(Table))` |
| Missing count | `COUNTROWS(FILTER(Table, Table[CompletionStatus]="Missing"))` |
| Filled by category | Slicer on `Category` + card with Completion % |
| Missing params list | Table visual filtered to `CompletionStatus = "Missing"` |

## Configuration

Edit the `CONFIG` dict at the top of [`revit_param_extractor.py`](revit_param_extractor.py):

```python
CONFIG = {
    "categories":          None,       # None = all, or ["OST_Walls", "OST_Doors"]
    "output_format":       "LONG",     # "LONG" or "WIDE"
    "output_dir":          None,       # None = file dialog prompt
    "parquet_compression": "snappy",   # "snappy" | "gzip" | "brotli" | None
    "skip_param_prefixes": [],
    "include_param_contains": [],
    "add_completion_flag": True,
}
```
