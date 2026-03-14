"""
Revit Parameter Extractor for Power BI
=======================================
Extracts all instance + type parameters from Revit elements
and writes a Parquet file optimised for Power BI
KPI dashboards and completion-status tracking.

Requires (CPython only — install once):
    pip install pyarrow pandas

Compatible with:
  - pyRevit  (CPython 3.x engine — set in pyRevit settings)
  - Dynamo   (put the script in a PythonScript node)
  - Revit API standalone

Output formats
  - LONG  : one row per (element × parameter)  ← best for Power BI slicers/KPIs
  - WIDE  : one row per element, columns = params ← best for quick tables

Usage (pyRevit button or script panel):
    Just run the file – a SaveFileDialog appears for the output path.
"""

# ---------------------------------------------------------------------------
# 0. Environment detection
# ---------------------------------------------------------------------------
import sys
import os
import datetime
import traceback

try:
    from collections import OrderedDict
except ImportError:
    OrderedDict = dict  # IronPython 2.6 fallback

# Revit API imports (available in pyRevit / Dynamo)
try:
    import clr
    clr.AddReference("RevitAPI")
    clr.AddReference("RevitAPIUI")
    from Autodesk.Revit.DB import (
        FilteredElementCollector,
        ElementFilter,
        BuiltInCategory,
        BuiltInParameter,
        StorageType,
        Element,
        ElementId,
        UnitUtils,
        ForgeTypeId,
        SpecTypeId,
    )
    try:
        # Revit 2022+ uses ForgeTypeId for units
        from Autodesk.Revit.DB import UnitTypeId
        REVIT_NEW_UNITS = True
    except ImportError:
        REVIT_NEW_UNITS = False

    # pyRevit exposes __revit__; Dynamo exposes IN[0] as doc
    try:
        doc   = __revit__.ActiveUIDocument.Document  # type: ignore
        uidoc = __revit__.ActiveUIDocument            # type: ignore
        HOST  = "pyrevit"
    except NameError:
        try:
            doc  = IN[0]   # type: ignore  # Dynamo
            HOST = "dynamo"
        except NameError:
            doc  = None
            HOST = "standalone"

    REVIT_AVAILABLE = True
except Exception:
    REVIT_AVAILABLE = False
    HOST = "standalone"


# ---------------------------------------------------------------------------
# 1. Configuration  ← edit these to match your project
# ---------------------------------------------------------------------------
CONFIG = {
    # ---- Which elements to export ----------------------------------------
    # Set to None to export ALL model elements (slow on large models).
    # Use a list of BuiltInCategory names (strings) to filter.
    # Example: ["OST_Walls", "OST_Doors", "OST_Windows", "OST_Floors"]
    "categories": None,          # None = all categories

    # ---- Output ----------------------------------------------------------
    "output_format":  "LONG",   # "LONG" or "WIDE"
    "output_dir":     None,     # None = prompt user; or set an absolute path
    "file_prefix":   "revit_params",
    # Parquet compression: "snappy" (fast), "gzip" (smaller), "brotli" (smallest), None
    "parquet_compression": "snappy",

    # ---- Filtering -------------------------------------------------------
    # Skip parameters whose name starts with any of these prefixes
    "skip_param_prefixes": [],

    # Include only parameters whose name contains one of these substrings
    # (empty list = include everything)
    "include_param_contains": [],

    # Skip elements that have no work-set or whose work-set name contains:
    "skip_worksets": [],

    # ---- Metadata columns always added to every row ----------------------
    "meta_fields": [
        "ElementId",
        "Category",
        "FamilyName",
        "TypeName",
        "LevelName",
        "WorksetName",
        "Phase",
    ],

    # ---- KPI helpers (LONG format only) ---------------------------------
    # If a parameter value is empty/None, mark completion as "Missing"
    # so Power BI can count filled vs unfilled parameters.
    "add_completion_flag": True,
}

# ---------------------------------------------------------------------------
# 2. Utility helpers
# ---------------------------------------------------------------------------

def safe_str(value):
    """Convert any value to a clean unicode string, never raises."""
    if value is None:
        return ""
    try:
        s = str(value)
        return s.strip()
    except Exception:
        return ""


def get_param_value(param):
    """
    Read a Revit Parameter and return (display_value, raw_value, unit_label).
    Handles all StorageType variants and Revit 2021/2022+ unit API changes.
    """
    if param is None or not param.HasValue:
        return ("", None, "")

    st = param.StorageType

    try:
        if st == StorageType.String:
            val = param.AsString() or ""
            return (val, val, "")

        elif st == StorageType.Integer:
            raw = param.AsInteger()
            # Boolean parameters
            try:
                if param.Definition.GetDataType() == SpecTypeId.Boolean.YesNo:
                    display = "Yes" if raw else "No"
                    return (display, raw, "")
            except Exception:
                pass
            return (safe_str(raw), raw, "")

        elif st == StorageType.Double:
            raw   = param.AsDouble()
            unit  = ""
            try:
                # Revit 2022+ ForgeTypeId units
                if REVIT_NEW_UNITS:
                    type_id  = param.GetUnitTypeId()
                    display  = UnitUtils.ConvertFromInternalUnits(raw, type_id)
                    unit     = safe_str(type_id)
                else:
                    display_str = param.AsValueString()
                    display     = display_str if display_str else raw
            except Exception:
                display = raw

            return (safe_str(display), raw, unit)

        elif st == StorageType.ElementId:
            eid    = param.AsElementId()
            raw_id = eid.IntegerValue if eid else -1
            # Try to resolve the element name for human-readable output
            try:
                linked_elem = doc.GetElement(eid)
                name = safe_str(linked_elem.Name) if linked_elem else safe_str(raw_id)
            except Exception:
                name = safe_str(raw_id)
            return (name, raw_id, "ElementId")

        else:
            return (safe_str(param.AsValueString()), None, "")

    except Exception as exc:
        return ("ERROR: " + safe_str(exc), None, "")


# ---------------------------------------------------------------------------
# 3. Element metadata helper
# ---------------------------------------------------------------------------

def get_element_meta(elem):
    """Return an OrderedDict with the metadata columns defined in CONFIG."""
    meta = OrderedDict()

    meta["ElementId"] = elem.Id.IntegerValue

    # Category
    try:
        meta["Category"] = safe_str(elem.Category.Name) if elem.Category else ""
    except Exception:
        meta["Category"] = ""

    # Family + Type
    try:
        sym  = doc.GetElement(elem.GetTypeId())
        if sym:
            meta["FamilyName"] = safe_str(sym.FamilyName)
            meta["TypeName"]   = safe_str(sym.Name)
        else:
            meta["FamilyName"] = ""
            meta["TypeName"]   = ""
    except Exception:
        meta["FamilyName"] = ""
        meta["TypeName"]   = ""

    # Level
    try:
        lvl_id = elem.LevelId
        level  = doc.GetElement(lvl_id)
        meta["LevelName"] = safe_str(level.Name) if level else ""
    except Exception:
        meta["LevelName"] = ""

    # Workset
    try:
        from Autodesk.Revit.DB import WorksharingUtils
        ws_info = WorksharingUtils.GetWorksharingTooltipInfo(doc, elem.Id)
        meta["WorksetName"] = safe_str(ws_info.WorksetName)
    except Exception:
        meta["WorksetName"] = ""

    # Phase (look for Phase Created parameter)
    try:
        phase_param = elem.get_Parameter(BuiltInParameter.PHASE_CREATED)
        if phase_param and phase_param.HasValue:
            phase_elem = doc.GetElement(phase_param.AsElementId())
            meta["Phase"] = safe_str(phase_elem.Name) if phase_elem else ""
        else:
            meta["Phase"] = ""
    except Exception:
        meta["Phase"] = ""

    return meta


# ---------------------------------------------------------------------------
# 4. Parameter collection for one element
# ---------------------------------------------------------------------------

def collect_element_params(elem):
    """
    Returns a list of dicts, one per parameter:
      {ParameterName, GroupName, IsInstance, StorageTypeName,
       DisplayValue, RawValue, Unit}
    """
    rows = []
    seen = set()  # de-duplicate by (name, group) when type params repeat

    # Instance parameters
    for param in elem.Parameters:
        try:
            name  = safe_str(param.Definition.Name)
            group = safe_str(param.Definition.ParameterGroup)
            key   = (name, group, True)
            if key in seen:
                continue
            seen.add(key)

            if _should_skip(name):
                continue

            disp, raw, unit = get_param_value(param)
            rows.append({
                "ParameterName":  name,
                "GroupName":      group,
                "IsInstance":     True,
                "StorageType":    safe_str(param.StorageType),
                "DisplayValue":   disp,
                "RawValue":       safe_str(raw),
                "Unit":           unit,
            })
        except Exception:
            continue

    # Type parameters (if element has a type)
    try:
        type_elem = doc.GetElement(elem.GetTypeId())
        if type_elem:
            for param in type_elem.Parameters:
                try:
                    name  = safe_str(param.Definition.Name)
                    group = safe_str(param.Definition.ParameterGroup)
                    key   = (name, group, False)
                    if key in seen:
                        continue
                    seen.add(key)

                    if _should_skip(name):
                        continue

                    disp, raw, unit = get_param_value(param)
                    rows.append({
                        "ParameterName":  name,
                        "GroupName":      group,
                        "IsInstance":     False,
                        "StorageType":    safe_str(param.StorageType),
                        "DisplayValue":   disp,
                        "RawValue":       safe_str(raw),
                        "Unit":           unit,
                    })
                except Exception:
                    continue
    except Exception:
        pass

    return rows


def _should_skip(param_name):
    """Apply CONFIG-based name filters."""
    name_lower = param_name.lower()
    for prefix in CONFIG["skip_param_prefixes"]:
        if name_lower.startswith(prefix.lower()):
            return True
    includes = CONFIG["include_param_contains"]
    if includes:
        if not any(s.lower() in name_lower for s in includes):
            return True
    return False


# ---------------------------------------------------------------------------
# 5. Main extraction loop
# ---------------------------------------------------------------------------

def collect_all_elements():
    """
    Yield every element that should be processed according to CONFIG.
    Uses FilteredElementCollector for speed.
    """
    if not REVIT_AVAILABLE or doc is None:
        return

    cats = CONFIG.get("categories")

    if cats:
        for cat_name in cats:
            try:
                bic = getattr(BuiltInCategory, cat_name)
                collector = (FilteredElementCollector(doc)
                             .OfCategory(bic)
                             .WhereElementIsNotElementType())
                for elem in collector:
                    if _workset_ok(elem):
                        yield elem
            except Exception as exc:
                print("Warning – category {}: {}".format(cat_name, exc))
    else:
        # All non-type elements; skip annotation/detail items for speed
        collector = (FilteredElementCollector(doc)
                     .WhereElementIsNotElementType())
        for elem in collector:
            try:
                if elem.Category and elem.Category.Name and _workset_ok(elem):
                    yield elem
            except Exception:
                continue


def _workset_ok(elem):
    skip_ws = CONFIG.get("skip_worksets", [])
    if not skip_ws:
        return True
    try:
        from Autodesk.Revit.DB import WorksharingUtils
        ws = WorksharingUtils.GetWorksharingTooltipInfo(doc, elem.Id).WorksetName
        return not any(s.lower() in ws.lower() for s in skip_ws)
    except Exception:
        return True


# ---------------------------------------------------------------------------
# 6. Build rows
# ---------------------------------------------------------------------------

def build_long_rows():
    """One row per (element, parameter) — best for Power BI Power Query."""
    rows = []
    meta_keys = CONFIG["meta_fields"]

    for elem in collect_all_elements():
        try:
            meta   = get_element_meta(elem)
            params = collect_element_params(elem)
            for p in params:
                row = OrderedDict()
                # Metadata first (always same columns)
                for k in meta_keys:
                    row[k] = meta.get(k, "")
                # Parameter data
                row.update(p)
                # Completion flag
                if CONFIG["add_completion_flag"]:
                    row["CompletionStatus"] = (
                        "Filled" if p["DisplayValue"] not in ("", None) else "Missing"
                    )
                rows.append(row)
        except Exception:
            continue

    return rows


def build_wide_rows():
    """One row per element, one column per unique parameter name."""
    # First pass: gather all parameter names (preserving order)
    param_names = []
    seen_names  = set()
    element_data = []   # list of (meta, param_rows)

    for elem in collect_all_elements():
        try:
            meta   = get_element_meta(elem)
            params = collect_element_params(elem)
            element_data.append((meta, params))
            for p in params:
                n = p["ParameterName"]
                if n not in seen_names:
                    param_names.append(n)
                    seen_names.add(n)
        except Exception:
            continue

    # Second pass: build flat rows
    meta_keys = CONFIG["meta_fields"]
    rows = []
    for meta, params in element_data:
        row = OrderedDict()
        for k in meta_keys:
            row[k] = meta.get(k, "")
        # Build a lookup for this element's params
        p_lookup = {p["ParameterName"]: p["DisplayValue"] for p in params}
        for n in param_names:
            row[n] = p_lookup.get(n, "")
        rows.append(row)

    return rows


# ---------------------------------------------------------------------------
# 7. Parquet writer
# ---------------------------------------------------------------------------

def _rows_to_dataframe(rows):
    """
    Convert list-of-dicts to a pandas DataFrame with correct dtypes.

    Column typing strategy (important for Power BI query folding):
      ElementId  → Int64  (nullable integer)
      IsInstance → bool
      RawValue   → float64 where possible, else string
      everything else → string (pandas StringDtype for nullable strings)
    """
    import pandas as pd

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    # --- ElementId: integer ---------------------------------------------------
    if "ElementId" in df.columns:
        df["ElementId"] = pd.to_numeric(df["ElementId"], errors="coerce").astype("Int64")

    # --- IsInstance: boolean --------------------------------------------------
    if "IsInstance" in df.columns:
        df["IsInstance"] = df["IsInstance"].astype(bool)

    # --- RawValue: try numeric, keep string on failure -----------------------
    if "RawValue" in df.columns:
        numeric = pd.to_numeric(df["RawValue"], errors="coerce")
        # Only convert if at least 30% of values parsed as numeric
        if numeric.notna().mean() >= 0.30:
            df["RawValue"] = numeric          # float64
        else:
            df["RawValue"] = df["RawValue"].astype(str).replace("None", pd.NA)

    # --- All remaining object columns → StringDtype (nullable) ---------------
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).replace({"None": pd.NA, "nan": pd.NA})
        df[col] = df[col].astype(pd.StringDtype())

    return df


def write_parquet(rows, filepath):
    """
    Write rows to a Parquet file.

    Tries pyarrow first (fastest, best Parquet support),
    falls back to pandas with fastparquet engine if pyarrow is missing.
    """
    if not rows:
        print("No rows to write.")
        return

    compression = CONFIG.get("parquet_compression", "snappy")

    # --- pyarrow path (preferred) --------------------------------------------
    try:
        import pyarrow as pa
        import pyarrow.parquet as pq
        import pandas as pd

        df    = _rows_to_dataframe(rows)
        table = pa.Table.from_pandas(df, preserve_index=False)
        pq.write_table(
            table,
            filepath,
            compression=compression,
            write_statistics=True,    # enables predicate pushdown in Power BI
            row_group_size=100_000,   # good for large models
        )
        size_mb = os.path.getsize(filepath) / 1_048_576
        print("Parquet written (pyarrow): {}  [{:.2f} MB, compression={}]".format(
            filepath, size_mb, compression))
        return

    except ImportError:
        pass

    # --- pandas + fastparquet fallback ---------------------------------------
    try:
        import pandas as pd
        df = _rows_to_dataframe(rows)
        df.to_parquet(filepath, engine="fastparquet", compression=compression, index=False)
        size_mb = os.path.getsize(filepath) / 1_048_576
        print("Parquet written (fastparquet): {}  [{:.2f} MB]".format(filepath, size_mb))
        return

    except ImportError:
        pass

    raise RuntimeError(
        "Neither pyarrow nor fastparquet is installed.\n"
        "Run:  pip install pyarrow pandas\n"
        "inside the CPython environment used by pyRevit."
    )


# ---------------------------------------------------------------------------
# 8. Output path resolution
# ---------------------------------------------------------------------------

def resolve_output_path():
    ts     = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = CONFIG["file_prefix"]
    fmt    = CONFIG["output_format"].lower()
    name   = "{}_{}_{}.parquet".format(prefix, fmt, ts)

    out_dir = CONFIG.get("output_dir")
    if out_dir:
        return os.path.join(out_dir, name)

    # Try to show a native Save dialog (works in pyRevit on Windows)
    try:
        import System.Windows.Forms as WinForms  # type: ignore
        dlg            = WinForms.SaveFileDialog()
        dlg.Title      = "Save Revit Parameter Export"
        dlg.FileName   = name
        dlg.DefaultExt = "parquet"
        dlg.Filter     = "Parquet files (*.parquet)|*.parquet|All files (*.*)|*.*"
        if str(dlg.ShowDialog()) == "OK":
            return dlg.FileName
    except Exception:
        pass

    # Fallback: Desktop
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        desktop = os.path.expanduser("~")
    return os.path.join(desktop, name)


# ---------------------------------------------------------------------------
# 9. Entry point
# ---------------------------------------------------------------------------

def run():
    if not REVIT_AVAILABLE or doc is None:
        print("ERROR: Revit document not available. Run this inside pyRevit or Dynamo.")
        return

    print("=" * 60)
    print("Revit Parameter Extractor  –  Power BI Export")
    print("Format : {}".format(CONFIG["output_format"]))
    print("=" * 60)

    t0 = datetime.datetime.now()

    if CONFIG["output_format"] == "LONG":
        rows = build_long_rows()
    else:
        rows = build_wide_rows()

    elapsed = (datetime.datetime.now() - t0).total_seconds()
    unique_elems = len(set(r.get("ElementId", "") for r in rows))
    print("Collected {:,} rows from {:,} elements in {:.1f}s".format(
        len(rows), unique_elems, elapsed))

    if not rows:
        print("No data collected. Check category filters and model content.")
        return

    path = resolve_output_path()
    write_parquet(rows, path)

    print("Done.")
    return rows   # also return for Dynamo OUT variable


# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------
if __name__ == "__main__" or HOST in ("pyrevit", "dynamo"):
    result = run()
    try:
        OUT = result   # Dynamo output port
    except Exception:
        pass
