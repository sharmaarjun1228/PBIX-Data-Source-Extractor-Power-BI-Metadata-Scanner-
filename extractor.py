import os
import re
from datetime import datetime

import pandas as pd
from pbixray import PBIXRay

# ---------- REGEX PATTERNS ----------
SQL_FROM_PATTERN = re.compile(r'\bFROM\s+([A-Za-z0-9_\[\]\.]+)', re.IGNORECASE)
SQL_JOIN_PATTERN = re.compile(r'\bJOIN\s+([A-Za-z0-9_\[\]\.]+)', re.IGNORECASE)
SQL_EXEC_PATTERN = re.compile(r'\bEXEC(?:UTE)?\s+([A-Za-z0-9_\[\]\.]+)', re.IGNORECASE)

SQL_DATABASE_PATTERN = re.compile(
    r'Sql\.Database\(\s*"([^"]+)"\s*,\s*"([^"]+)"',
    re.IGNORECASE
)

NATIVE_QUERY_PATTERN = re.compile(
    r'Value\.NativeQuery\([^,]+,\s*"(.+?)"\s*(?:,|\))',
    re.IGNORECASE | re.DOTALL
)

QUERY_OPTION_PATTERN = re.compile(
    r'Query\s*=\s*"(.+?)"',
    re.IGNORECASE | re.DOTALL
)

SCHEMA_ITEM_PATTERN = re.compile(
    r'Schema\s*=\s*"([^"]+)"\s*,\s*Item\s*=\s*"([^"]+)"',
    re.IGNORECASE
)

ITEM_ONLY_PATTERN = re.compile(
    r'Item\s*=\s*"([^"]+)"',
    re.IGNORECASE
)
# -------------------------------------


def extract_objects_from_sql(sql: str):
    objs = set()
    for pattern in (SQL_FROM_PATTERN, SQL_JOIN_PATTERN, SQL_EXEC_PATTERN):
        for match in pattern.findall(sql):
            obj = match.strip()
            if len(obj) >= 3:
                objs.add(obj)
    return sorted(objs)


def process_pbix_file(pbix_path: str):
    rows = []

    print(f"\n=== Processing {pbix_path} ===")
    try:
        model = PBIXRay(pbix_path)
        pq_df = model.power_query
        if pq_df is None or pq_df.empty:
            print("  (No Power Query / M code found in this file.)")
            return rows
    except Exception as e:
        print(f"  !! Error reading {pbix_path}: {e}")
        rows.append({
            "pbix_file": pbix_path,
            "table_name": None,
            "source_type": "ERROR",
            "server": None,
            "database": None,
            "object_name": None,
            "sql_snippet": str(e)
        })
        return rows

    for _, rec in pq_df.iterrows():
        table_name = rec.get("TableName") or rec.get("Name")
        expression = str(rec.get("Expression") or "")

        server = None
        database = None
        m_conn = SQL_DATABASE_PATTERN.search(expression)
        if m_conn:
            server, database = m_conn.group(1), m_conn.group(2)

        native_match = NATIVE_QUERY_PATTERN.search(expression)
        query_match = QUERY_OPTION_PATTERN.search(expression)
        schema_items = SCHEMA_ITEM_PATTERN.findall(expression)
        item_only = ITEM_ONLY_PATTERN.findall(expression)

        if native_match:
            sql = native_match.group(1)
            objects = extract_objects_from_sql(sql) or [None]
            for obj in objects:
                rows.append({
                    "pbix_file": pbix_path,
                    "table_name": table_name,
                    "source_type": "SQL (NativeQuery)",
                    "server": server,
                    "database": database,
                    "object_name": obj,
                    "sql_snippet": sql[:500]
                })

        if query_match:
            sql_q = query_match.group(1).replace('""', '"')
            objects = extract_objects_from_sql(sql_q) or [None]
            for obj in objects:
                rows.append({
                    "pbix_file": pbix_path,
                    "table_name": table_name,
                    "source_type": "SQL (Query option)",
                    "server": server,
                    "database": database,
                    "object_name": obj,
                    "sql_snippet": sql_q[:500]
                })

        for schema, item in schema_items:
            full_name = f"{schema}.{item}"
            rows.append({
                "pbix_file": pbix_path,
                "table_name": table_name,
                "source_type": "Navigation (Schema/Item)",
                "server": server,
                "database": database,
                "object_name": full_name,
                "sql_snippet": None
            })

        if not schema_items and item_only:
            for item in item_only:
                rows.append({
                    "pbix_file": pbix_path,
                    "table_name": table_name,
                    "source_type": "Navigation (Item only)",
                    "server": server,
                    "database": database,
                    "object_name": item,
                    "sql_snippet": None
                })

        if not (native_match or query_match or schema_items or item_only):
            if server or database:
                rows.append({
                    "pbix_file": pbix_path,
                    "table_name": table_name,
                    "source_type": "M query (connection only)",
                    "server": server,
                    "database": database,
                    "object_name": None,
                    "sql_snippet": None
                })

    return rows


def classify_object(name: str):
    if not name:
        return None
    n = name.lower()
    if "." in n:
        n = n.split(".")[-1]

    if "usp" in n or n.startswith("sp_"):
        return "Stored Procedure"
    if "vw" in n or "view" in n:
        return "View"
    return "Table/View (unknown)"


def run_extraction(base_folder: str) -> str:
    """
    Runs the extraction on all PBIX files under base_folder.
    Returns the path to the created Excel file.
    """
    if not os.path.isdir(base_folder):
        raise ValueError(f"Folder does not exist: {base_folder}")

    all_rows = []
    pbix_count = 0

    print("\nScanning for PBIX files...\n")
    for root, dirs, files in os.walk(base_folder):
        for filename in files:
            if filename.lower().endswith(".pbix"):
                pbix_count += 1
                full_path = os.path.join(root, filename)
                print(f"  -> Found PBIX: {full_path}")
                file_rows = process_pbix_file(full_path)
                all_rows.extend(file_rows)

    if pbix_count == 0:
        raise RuntimeError("No .pbix files found under this folder.")

    if not all_rows:
        raise RuntimeError("PBIX files found but no data sources extracted.")

    df = pd.DataFrame(all_rows)
    df["object_type"] = df["object_name"].apply(classify_object)

    df.sort_values(
        by=["server", "database", "object_type", "object_name", "pbix_file"],
        inplace=True,
        na_position="last"
    )

    summary = (
        df.dropna(subset=["object_name"])
          .groupby(["server", "database", "object_name", "object_type"], dropna=False)
          .agg(
              pbix_files=("pbix_file", lambda x: "; ".join(sorted({
                  os.path.basename(p) for p in x
              })))
          )
          .reset_index()
    )

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(base_folder, f"pbix_data_sources_{ts}.xlsx")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Detailed", index=False)
        summary.to_excel(writer, sheet_name="UniqueObjects", index=False)

    print(f"\n✅ Done! Excel created at: {output_file}")
    return output_file
