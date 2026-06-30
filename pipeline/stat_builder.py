import os
import re
import json
import pandas as pd
import pyreadstat
from io import BytesIO
from datetime import datetime, date

from schema_builder import build_rawdata_spec

# =========================================================
# 固定欄位（避免空 dataframe 沒欄位時 merge 出錯）
# =========================================================
SCHEMA_COLUMNS = [
    "SOURCE",
    "FORM",
    "RAW_DATASET",
    "STAT_DATASET",
    "VARIABLE_ORDER",
    "VARIABLE",
    "LABEL",
    "SCHEMA_DATATYPE",
    "SCHEMA_DATATYPE_STD",
]

RAW_COLUMNS = [
    "SOURCE",
    "RAW_DATASET",
    "VARIABLE",
    "RAW_DATATYPE",
    "RAW_DATATYPE_STD",
    "RAW_SAMPLE",
]

COMPARE_COLUMNS = [
    "FORM",
    "RAW_DATASET",
    "STAT_DATASET",
    "VARIABLE_ORDER",
    "VARIABLE",
    "LABEL",
    "SCHEMA_DATATYPE",
    "SCHEMA_DATATYPE_STD",
    "RAW_DATATYPE",
    "RAW_DATATYPE_STD",
    "RAW_SAMPLE",
    "STATUS",
]



# =================================================================================================================
# 文字處理
# =================================================================================================================
def normalize_text(x):

    if pd.isna(x):
        return ""
    x = str(x) #統一資料型態
    x = x.replace("\n", " ").replace("\r", " ").replace("\xa0", " ") #移除換行
    x = re.sub(r"\s*,\s*", ", ", x) #統一逗號格式
    x = re.sub(r"\s+", " ", x) #壓縮多於空白

    return x.strip().upper()



# =========================================================
# datatype normalize
# =========================================================
def normalize_raw_dtype(dtype):
    d = str(dtype).lower()

    if "datetime" in d:
        return "DATETIME"
    elif "date" in d:
        return "DATE"
    elif "int" in d or "float" in d or "double" in d:
        return "NUMERIC"
    else:
        return "TEXT"




def match_folder_visit(raw_visit: str, folder_full_term: str, repeat_folder, abbr):
    """
    回傳 match type:
      - EXACT
      - ALL_VISIT_FUZZY
      - REPEAT_FUZZY
      - None
    """
    raw_norm = normalize_text(raw_visit)
    folder_norm = normalize_text(folder_full_term)

    # 1) exact after normalization
    if raw_norm == folder_norm:
        return "EXACT"


    # 2) All Visit - XXX
    if raw_norm.startswith("ALL VISIT - "):
        tail = re.sub(r"^\s*ALL\s*VISIT\s*-\s*", "", raw_norm).strip()

        if tail == folder_norm:
            return "ALL_VISIT_EXACT"

        if tail in folder_norm or folder_norm in tail:
            return "ALL_VISIT_FUZZY"



    # 3) 去尾巴數字
    raw_no_num = re.sub(r"\s+\d+$", "", raw_norm)

    if raw_no_num == folder_norm:
        return "TAIL_NUMBER_FUZZY"


    # 4) repeat fuzzy
    repeat_flag = str(repeat_folder).strip().lower() in ["yes", "y", "true", "1"]

    if repeat_flag:
        # 直接在這裡抓 Day number，不另外拆 function
        raw_day_match = re.search(r"day\s+(\d+)", raw_norm, flags=re.IGNORECASE)
        folder_day_match = re.search(r"day\s+(\d+)", folder_norm, flags=re.IGNORECASE)

        raw_day = raw_day_match.group(1) if raw_day_match else None
        folder_day = folder_day_match.group(1) if folder_day_match else None

        # 例：
        # Folder: Additional Cycle Day 1
        # Raw:    Cycle 3, Day 1

        if raw_day and folder_day and raw_day == folder_day:
            return "REPEAT_CYCLE_FUZZY"

        # UN1 / UN2 / UN3
        if re.fullmatch(fr"{abbr}\d+", raw_norm):
            return "REPEAT_ABBR_FUZZY"

    return None







def infer_visitnum_from_raw_visit(raw_visit: str):
    """
    自動推導 VISITNUM
    回傳 int 或 None
    """

    s = normalize_text(raw_visit)

    # ---------------------------------
    # 1) Screening / Baseline / EOS / Follow-up
    # ---------------------------------
    if "SCREEN" in s or "ENROLLMENT" in s:
        return -900

    if "BASELINE" in s:
        return 0

    if "END OF STUDY" in s or "EOS" == s:
        return 999000

    if "END OF TREATMENT" in s or "EOT" == s:
        return 998000

    if "FOLLOW-UP" in s or "FOLLOW UP" in s:
        return 999900
    
    if "ADVERSE EVENT" in s or "CONCOMITANT" in s:
        return 1000000

    # ---------------------------------
    # 2) Cycle N, Day M
    #   e.g. Cycle 3, Day 1 -> 301
    # ---------------------------------
    m = re.search(r"CYCLE\s+(\d+)\s*,?\s*DAY\s+(\d+)", s)
    if m:
        cycle_no = int(m.group(1))
        day_no = int(m.group(2))
        return cycle_no * 10000 + day_no * 100

    # ---------------------------------
    # 3) Period N Day M
    #   e.g. Period 2 Day 1 -> 2001
    #   我故意跟 Cycle 區分開，不然排序會撞
    # ---------------------------------
    m = re.search(r"PERIOD\s+(\d+)\s*DAY\s+(\d+)", s)
    if m:
        period_no = int(m.group(1))
        day_no = int(m.group(2))
        return period_no * 10000 + day_no * 100

    # ---------------------------------
    # 4) Visit N (Day M)
    #   e.g. Visit 2 (Day 10) -> 210
    # ---------------------------------
    m = re.search(r"VISIT\s+(\d+)\s*\(\s*DAY\s+(\d+)\s*\)", s)
    if m:
        visit_no = int(m.group(1))
        day_no = int(m.group(2))
        return visit_no * 100

    # ---------------------------------
    # 5) Visit N
    #   e.g. Visit 1 -> 100, Visit 2 -> 200
    # ---------------------------------
    m = re.search(r"VISIT\s+(\d+)$", s)
    if m:
        visit_no = int(m.group(1))
        return visit_no * 100

    # ---------------------------------
    # 6) Week N
    #   e.g. Week 1 -> 7001, Week 2 -> 7002
    #   用 7000+N 避免和 Visit/Cycle 混淆
    # ---------------------------------
    m = re.search(r"WEEK\s+(\d+)$", s)
    if m:
        week_no = int(m.group(1))
        return week_no * 100

    # ---------------------------------
    # 7) Unscheduled
    #   e.g. UN1 -> 9001
    # ---------------------------------
    m = re.fullmatch(r"UN\s*(\d+)", s)
    if m:
        return 999999

    # ---------------------------------
    # 8) Tumor Assessment N
    #   e.g. Tumor Assessment 1 -> 8001
    # ---------------------------------
    m = re.search(r"TUMOR ASSESSMENT\s+(\d+)$", s)
    if m:
        return 900000

    return None





def build_visit_mapping_table(folder_map_df: pd.DataFrame, raw_dataset_map: dict) -> pd.DataFrame:
    """
    建立 raw VISIT 與 Folder FULL_TERM 的對照表
    """

    # 保底
    if folder_map_df is None:
        folder_map_df = pd.DataFrame(columns=["VISIT", "VISIT_DVP", "VISITNUM", "REPEAT_FOLDER"])

    if raw_dataset_map is None:
        raw_dataset_map = {}


    # =================================================
    # 1) 收集所有 raw VISIT 值
    # =================================================
    raw_values = set()

    for ds_name, df in raw_dataset_map.items():
        if df is None or df.empty:
            continue

        cols_upper = {str(c).strip().upper(): c for c in df.columns}

        if "VISIT" not in cols_upper:
            continue

        real_col = cols_upper["VISIT"]

        vals = (
            df[real_col]
            .dropna()
            .astype(str)
            .str.strip()
        )

        vals = vals[vals != ""]
        raw_values.update(vals.tolist())

    # =================================================
    # 2) Folder rows normalize
    # =================================================
    folder_rows = folder_map_df.copy()

    if folder_rows.empty:
        # 沒有 Folder，只能列 raw-only
        out_rows = []
        for raw_visit in sorted(raw_values):

            visitnum = infer_visitnum_from_raw_visit(raw_visit)

            out_rows.append({
                "SOURCE": "RAWDATA_ONLY",
                "VISIT": raw_visit,
                "RAW_VISIT_NORMALIZED": normalize_text(raw_visit),
                "MATCH_TYPE": "UNMATCHED",
                "VISIT_DVP": "",
                "FULL_TERM": "",
                "VISITNUM": visitnum if visitnum is not None else pd.NA,
                "REPEAT_FOLDER": ""
            })
        return pd.DataFrame(out_rows)


    folder_rows["VISIT_DVP"] = folder_rows["VISIT_DVP"].astype(str).str.strip()

    if "VISITNUM" not in folder_rows.columns:
        folder_rows["VISITNUM"] = pd.NA

    if "REPEAT_FOLDER" not in folder_rows.columns:
        folder_rows["REPEAT_FOLDER"] = pd.NA

    out_rows = []

    # =================================================
    # 3) raw VISIT vs Folder FULL_TERM matching
    # =================================================
    matched_folder_keys = set()

    for raw_visit in sorted(raw_values):

        raw_norm = normalize_text(raw_visit)

        matched_abbr = ""
        matched_full = ""
        matched_visitnum = ""
        matched_repeat = ""
        match_type = "UNMATCHED"
        matched_idx = None

        for idx, r in folder_rows.iterrows():

            mtype = match_folder_visit(
                raw_visit=raw_visit,
                folder_full_term=r["VISIT"],
                repeat_folder=r["REPEAT_FOLDER"],
                abbr=r["VISIT_DVP"]
            )

            if mtype is not None:
                matched_abbr = r["VISIT_DVP"]
                matched_full = r["VISIT"]
                matched_repeat = r["REPEAT_FOLDER"]
                match_type = mtype
                matched_idx = idx
                break

        visitnum = infer_visitnum_from_raw_visit(raw_visit)

        if matched_idx is not None:
            matched_folder_keys.add(matched_idx)
            source_type = "MATCHED"
        else:
            source_type = "RAWDATA_ONLY"

        out_rows.append({
            "SOURCE": source_type,
            "VISIT": raw_visit,
            "RAW_VISIT_NORMALIZED": raw_norm,
            "MATCH_TYPE": match_type,
            "VISIT_DVP": matched_abbr,
            "FULL_TERM": matched_full,
            "VISITNUM": visitnum,
            "REPEAT_FOLDER": matched_repeat
        })



    # -----------------------------
    # 再補上 Folder-only rows（raw 尚未出現）
    # -----------------------------
    for idx, r in folder_rows.iterrows():
        if idx in matched_folder_keys:
            continue

        out_rows.append({
            "SOURCE": "SCHEMA_ONLY",
            "VISIT": "",
            "RAW_VISIT_NORMALIZED": "",
            "MATCH_TYPE": "NOT_IN_RAW",
            "VISIT_DVP": r["VISIT_DVP"],
            "FULL_TERM": r["VISIT"],
            "VISITNUM": pd.NA,
            "REPEAT_FOLDER": r["REPEAT_FOLDER"]
        })

    out_df = pd.DataFrame(out_rows)

    # 排序：MATCHED → SCHEMA_ONLY → RAWDATA_ONLY
    source_order = {
        "MATCHED": 0,
        "SCHEMA_ONLY": 1,
        "RAWDATA_ONLY": 2
    }

    out_df["_SORT"] = out_df["SOURCE"].map(source_order).fillna(9)
    out_df = out_df.sort_values(
        by=["_SORT", "VISIT_DVP", "VISIT"],
        na_position="last"
    ).drop(columns=["_SORT"]).reset_index(drop=True)

    return out_df




def load_sas_from_zip(zip_file):
    """
    讀取 ZIP 裡所有 .sas7bdat → 回傳 raw_dataset_map

    Output:
        dict[
            DATASET_NAME (str) : pd.DataFrame
        ]
    """

    import zipfile
    import tempfile
    import os
    import pyreadstat

    raw_dataset_map = {}

    try:
        zip_bytes = zip_file.getvalue()

        with tempfile.TemporaryDirectory() as tmpdir:

            # ✅ 寫入 ZIP
            zip_path = os.path.join(tmpdir, "data.zip")
            with open(zip_path, "wb") as f:
                f.write(zip_bytes)

            # ✅ 解壓
            with zipfile.ZipFile(zip_path, "r") as z:
                z.extractall(tmpdir)

            # ✅ 遍歷讀 SAS
            for root, _, files in os.walk(tmpdir):
                for file in files:

                    if file.lower().endswith(".sas7bdat"):

                        file_path = os.path.join(root, file)

                        try:
                            df, meta = pyreadstat.read_sas7bdat(file_path)

                            dataset_name = file.replace(".sas7bdat", "").upper()

                            raw_dataset_map[dataset_name] = df

                        except Exception as e:
                            print(f"[SAS LOAD ERROR] {file}: {e}")

    except Exception as e:
        print(f"[ZIP ERROR] {e}")

    return raw_dataset_map




# =========================================================
# 建立 RAW metadata + preview
# =========================================================
def build_raw_metadata_and_preview(raw_dataset_map, preview_n=20, sample_n=5):
    """
    Input:
        raw_dataset_map: dict[str, pd.DataFrame]

    Output:
        raw_df
        raw_preview_map
        raw_dataset_map（整理後）
    
    規則：
    - VAR_FMT 優先
    - logical variable 統一
    """

    raw_preview_map = {}
    cleaned_dataset_map = {}

    if raw_dataset_map is None or len(raw_dataset_map) == 0:
        return pd.DataFrame(columns=RAW_COLUMNS), raw_preview_map, cleaned_dataset_map

    meta_records = []

    for dataset, df in raw_dataset_map.items():

        if df is None or df.empty:
            continue

        dataset = str(dataset).strip().upper()

        df = df.copy()
        df.columns = [str(c).strip().upper() for c in df.columns]

        columns = list(df.columns)
        col_set = set(columns)

        logical_cols = []

        # =========================
        # base + _FMT 優先
        # =========================
        for col in columns:

            if col.endswith("_FMT"):
                continue

            fmt_col = f"{col}_FMT"

            if fmt_col in col_set:
                logical_cols.append((col, fmt_col))
            else:
                logical_cols.append((col, col))

        # =========================
        # orphan _FMT
        # =========================
        for col in columns:

            if col.endswith("_FMT"):
                base = col[:-4]

                if base not in col_set:
                    logical_cols.append((base, col))

        # =========================
        # build logical dataframe
        # =========================
        data_dict = {}

        for logical_var, source_col in logical_cols:
            data_dict[logical_var] = df[source_col]

        logical_df = pd.DataFrame(data_dict)

        # ✅ preview
        raw_preview_map[dataset] = logical_df.head(preview_n).copy()

        # ✅ full dataset
        cleaned_dataset_map[dataset] = logical_df.copy()

        # =========================
        # metadata
        # =========================
        for logical_var, source_col in logical_cols:

            series = df[source_col]

            sample_values = (
                series.dropna()
                .astype(str)
                .head(sample_n)
                .tolist()
            )

            meta_records.append({
                "SOURCE": "RAW",
                "RAW_DATASET": dataset,
                "VARIABLE": logical_var.strip().upper(),
                "RAW_DATATYPE": str(series.dtype),
                "RAW_DATATYPE_STD": normalize_raw_dtype(series.dtype),
                "RAW_SAMPLE": json.dumps(sample_values, ensure_ascii=False)
            })

    if not meta_records:
        return pd.DataFrame(columns=RAW_COLUMNS), raw_preview_map, cleaned_dataset_map

    raw_df = pd.DataFrame(meta_records, columns=RAW_COLUMNS)

    return raw_df, raw_preview_map, cleaned_dataset_map





# =========================================================
# 比較 schema vs raw
# =========================================================

def compare_schema_raw(schema_df, raw_df, system_vars=None):

    if system_vars is None:
        system_vars = []

    system_vars = [str(x).strip().upper() for x in system_vars]

    # =========================================================
    # Schema normalize
    # =========================================================
    if schema_df is None or schema_df.empty:
        schema_df = pd.DataFrame(columns=SCHEMA_COLUMNS)
    else:
        schema_df = schema_df.copy()

        if "RAW_DATASET" not in schema_df.columns:
            schema_df["RAW_DATASET"] = pd.NA

        if "STAT_DATASET" not in schema_df.columns:
            schema_df["STAT_DATASET"] = schema_df["RAW_DATASET"]

        for c in SCHEMA_COLUMNS:
            if c not in schema_df.columns:
                schema_df[c] = pd.NA

        schema_df = schema_df[SCHEMA_COLUMNS]

        schema_df["FORM"] = schema_df["FORM"].astype(str).str.strip().str.upper()
        schema_df["RAW_DATASET"] = schema_df["RAW_DATASET"].astype(str).str.strip().str.upper()
        schema_df["STAT_DATASET"] = schema_df["STAT_DATASET"].astype(str).str.strip().str.upper()
        schema_df["VARIABLE"] = schema_df["VARIABLE"].astype(str).str.strip().str.upper()

    # =========================================================
    # Raw normalize
    # =========================================================
    if raw_df is None or raw_df.empty:
        raw_df = pd.DataFrame(columns=RAW_COLUMNS)
    else:
        raw_df = raw_df.copy()

        for c in RAW_COLUMNS:
            if c not in raw_df.columns:
                raw_df[c] = pd.NA

        raw_df = raw_df[RAW_COLUMNS]

        raw_df["RAW_DATASET"] = raw_df["RAW_DATASET"].astype(str).str.strip().str.upper()
        raw_df["VARIABLE"] = raw_df["VARIABLE"].astype(str).str.strip().str.upper()

    # =========================================================
    # Compare by RAW_DATASET + VARIABLE
    # =========================================================
    compare_df = pd.merge(
        schema_df,
        raw_df,
        on=["RAW_DATASET", "VARIABLE"],
        how="outer",
        suffixes=("", "_RAW")
    )

    COMPATIBLE_RULES = {
        ("DATE", "TEXT"),
        ("TIME", "TEXT"),
        ("DATETIME", "TEXT"),
    }

    def get_status(row):

        var = str(row.get("VARIABLE", "")).upper()
        schema_dtype = row.get("SCHEMA_DATATYPE_STD")
        raw_dtype = row.get("RAW_DATATYPE_STD")

        schema_missing = pd.isna(row.get("SCHEMA_DATATYPE"))
        raw_missing = pd.isna(row.get("RAW_DATATYPE"))

        if schema_missing and not raw_missing and var in system_vars:
            return "SYSTEM_VAR"

        if schema_missing and not raw_missing:
            return "EXTRA_IN_RAW"

        if raw_missing and not schema_missing:
            return "MISSING_IN_RAW"

        if schema_missing and raw_missing:
            return "UNKNOWN"

        if schema_dtype == raw_dtype:
            return "MATCH"

        if (schema_dtype, raw_dtype) in COMPATIBLE_RULES:
            return "MATCH"

        return "DATATYPE_MISMATCH"

    compare_df["STATUS"] = compare_df.apply(get_status, axis=1)

    for c in COMPARE_COLUMNS:
        if c not in compare_df.columns:
            compare_df[c] = pd.NA

    return compare_df[COMPARE_COLUMNS]






# =========================
# 合併資料 (水平/垂直)
# =========================

def merge_subitems(schema_df, raw_dataset_map, merge_config=None):
    """
    多階段 merge：
    1. 同一 FORM 底下，先依 non-system variable overlap 分群
    2. 群內若 >1 sub-item → vertical merge
    3. 群與群之間 → horizontal merge

    Returns
    -------
    merged_dataset_map : dict[str, pd.DataFrame]
        key = FORM / STAT_DATASET-like merged result
    merge_log_df : pd.DataFrame
    """
    import pandas as pd
    from functools import reduce

    if merge_config is None:
        merge_config = {
            "enabled": True,
            "default_key_vars": ["SUBJID", "USUBJID", "VISIT", "VISITNUM", "SITEID", "STUDYID"],
            "form_rules": {}
        }

    def _normalize_cols(df):
        out = df.copy()
        out.columns = [str(c).strip().upper() for c in out.columns]
        return out

    def _common_keys(dfs, candidate_keys):
        candidate_keys = [str(x).strip().upper() for x in candidate_keys]
        return [k for k in candidate_keys if all(k in df.columns for df in dfs)]

    def _business_cols(df, key_vars):
        key_vars = [str(x).strip().upper() for x in key_vars]
        return set([
            c for c in df.columns
            if c not in key_vars
            and c not in ["SOURCEDN", "RAW_DATASET", "STAT_DATASET"]
        ])

    def _build_overlap_groups(subitem_dfs, key_vars):
        n = len(subitem_dfs)
        if n == 0:
            return []

        biz_sets = []
        for name, stat_ds, df in subitem_dfs:
            biz_sets.append(_business_cols(df, key_vars))

        graph = {i: set() for i in range(n)}

        for i in range(n):
            for j in range(i + 1, n):
                overlap = biz_sets[i].intersection(biz_sets[j])
                if overlap:
                    graph[i].add(j)
                    graph[j].add(i)

        visited = set()
        groups = []

        for i in range(n):
            if i in visited:
                continue

            stack = [i]
            comp = []

            while stack:
                node = stack.pop()
                if node in visited:
                    continue

                visited.add(node)
                comp.append(subitem_dfs[node])

                for nei in graph[node]:
                    if nei not in visited:
                        stack.append(nei)

            groups.append(comp)

        return groups

    def _vertical_merge_group(group, form):
        tmp = []
        names = []

        for raw_ds, stat_ds, df in group:
            part = df.copy()
            part["SOURCEDN"] = raw_ds
            part["RAW_DATASET"] = raw_ds
            part["STAT_DATASET"] = stat_ds
            tmp.append(part)
            names.append(raw_ds)

        merged = pd.concat(tmp, axis=0, ignore_index=True, sort=False)
        return merged, names

    def _single_group(group):
        raw_ds, stat_ds, df = group[0]
        out = df.copy()

        if "SOURCEDN" in out.columns:
            out = out.drop(columns=["SOURCEDN"])

        out["RAW_DATASET"] = raw_ds
        out["STAT_DATASET"] = stat_ds

        return out, [raw_ds]

    merged_dataset_map = {}
    merge_logs = []

    if schema_df is None or schema_df.empty:
        schema_df = pd.DataFrame(columns=["FORM", "RAW_DATASET", "STAT_DATASET", "VARIABLE"])

    schema_df = schema_df.copy()

    schema_df["FORM"] = schema_df["FORM"].astype(str).str.strip().str.upper()
    schema_df["RAW_DATASET"] = schema_df["RAW_DATASET"].astype(str).str.strip().str.upper()

    if "STAT_DATASET" not in schema_df.columns:
        schema_df["STAT_DATASET"] = schema_df["RAW_DATASET"]

    schema_df["STAT_DATASET"] = schema_df["STAT_DATASET"].astype(str).str.strip().str.upper()

    # =========================================================
    # Merge disabled
    # =========================================================
    if not merge_config.get("enabled", False):

        for ds, df in raw_dataset_map.items():
            raw_ds = str(ds).strip().upper()
            tmp = _normalize_cols(df)
            tmp["RAW_DATASET"] = raw_ds
            tmp["STAT_DATASET"] = raw_ds
            merged_dataset_map[raw_ds] = tmp

        merge_log_df = pd.DataFrame(columns=[
            "FORM", "RAW_DATASETS", "STAT_DATASET", "SUBITEMS",
            "MERGE_ENABLED", "MERGE_MODE", "KEY_VARS",
            "RESULT_DATASET", "NOTE"
        ])
        return merged_dataset_map, merge_log_df

    # =========================================================
    # FORM -> RAW_DATASET / STAT_DATASET
    # =========================================================
    form_map = (
        schema_df[["FORM", "RAW_DATASET", "STAT_DATASET"]]
        .dropna()
        .drop_duplicates()
        .groupby("FORM")
        .apply(lambda g: g[["RAW_DATASET", "STAT_DATASET"]].drop_duplicates().to_dict("records"))
        .to_dict()
    )

    default_key_vars = [
        str(x).strip().upper()
        for x in merge_config.get("default_key_vars", ["SUBJID", "USUBJID", "VISIT", "VISITNUM"])
    ]

    all_used_subitems = set()

    for form, item_records in form_map.items():

        form_rule = merge_config.get("form_rules", {}).get(form, {})
        enabled = form_rule.get("enabled", True)
        key_vars = [
            str(x).strip().upper()
            for x in form_rule.get("key_vars", default_key_vars)
        ]

        existing = []

        for rec in item_records:
            raw_ds = str(rec["RAW_DATASET"]).strip().upper()
            stat_ds = str(rec["STAT_DATASET"]).strip().upper()

            if raw_ds in raw_dataset_map:
                df = _normalize_cols(raw_dataset_map[raw_ds])
                df["RAW_DATASET"] = raw_ds
                df["STAT_DATASET"] = stat_ds
                existing.append((raw_ds, stat_ds, df))

        if len(existing) == 0:
            merge_logs.append({
                "FORM": form,
                "RAW_DATASETS": "",
                "STAT_DATASET": "",
                "SUBITEMS": "",
                "MERGE_ENABLED": enabled,
                "MERGE_MODE": "skip",
                "KEY_VARS": ", ".join(key_vars),
                "RESULT_DATASET": "",
                "NOTE": "no raw sub-item dataset found"
            })
            continue

        for raw_ds, _, _ in existing:
            all_used_subitems.add(raw_ds)

        # 此 FORM 的主要 STAT_DATASET
        stat_dataset = (
            schema_df.loc[schema_df["FORM"] == form, "STAT_DATASET"]
            .dropna()
            .astype(str)
            .str.upper()
            .mode()
        )
        result_dataset = stat_dataset.iloc[0] if len(stat_dataset) else form

        if not enabled:
            for raw_ds, stat_ds, df in existing:
                merged_dataset_map[raw_ds] = df.copy()

            merge_logs.append({
                "FORM": form,
                "RAW_DATASETS": ", ".join([x[0] for x in existing]),
                "STAT_DATASET": result_dataset,
                "SUBITEMS": ", ".join([x[0] for x in existing]),
                "MERGE_ENABLED": enabled,
                "MERGE_MODE": "skip",
                "KEY_VARS": ", ".join(key_vars),
                "RESULT_DATASET": ", ".join([x[0] for x in existing]),
                "NOTE": "merge disabled by rule"
            })
            continue

        # 只有一個 sub-item
        if len(existing) == 1:
            single_df, single_names = _single_group(existing)
            merged_dataset_map[result_dataset] = single_df

            merge_logs.append({
                "FORM": form,
                "RAW_DATASETS": ", ".join(single_names),
                "STAT_DATASET": result_dataset,
                "SUBITEMS": ", ".join(single_names),
                "MERGE_ENABLED": enabled,
                "MERGE_MODE": "single",
                "KEY_VARS": ", ".join(key_vars),
                "RESULT_DATASET": result_dataset,
                "NOTE": "only one sub-item"
            })
            continue

        # =====================================================
        # 1) 先依 overlap 分群
        # =====================================================
        groups = _build_overlap_groups(existing, key_vars)

        group_results = []

        for idx, group in enumerate(groups, start=1):

            if len(group) == 1:
                g_df, g_names = _single_group(group)
                g_mode = "single"
            else:
                g_df, g_names = _vertical_merge_group(group, form=form)
                g_mode = "vertical"

            group_name = f"{form}__GRP{idx}"
            group_results.append((group_name, g_df, g_mode, g_names))

        # =====================================================
        # 2) 只有一個 group
        # =====================================================
        if len(group_results) == 1:

            _, final_df, final_mode, final_names = group_results[0]
            merged_dataset_map[result_dataset] = final_df

            merge_logs.append({
                "FORM": form,
                "RAW_DATASETS": ", ".join(final_names),
                "STAT_DATASET": result_dataset,
                "SUBITEMS": ", ".join(final_names),
                "MERGE_ENABLED": enabled,
                "MERGE_MODE": final_mode,
                "KEY_VARS": ", ".join(key_vars),
                "RESULT_DATASET": result_dataset,
                "NOTE": f"single group ({final_mode})"
            })
            continue

        # =====================================================
        # 3) 多 group → horizontal merge
        # =====================================================
        mergeable_dfs = [g[1] for g in group_results]
        common_keys = _common_keys(mergeable_dfs, key_vars)

        if len(common_keys) == 0:

            for group_name, g_df, g_mode, g_names in group_results:
                out_name = f"{result_dataset}__{group_name}"
                merged_dataset_map[out_name] = g_df

                merge_logs.append({
                    "FORM": form,
                    "RAW_DATASETS": ", ".join(g_names),
                    "STAT_DATASET": result_dataset,
                    "SUBITEMS": ", ".join(g_names),
                    "MERGE_ENABLED": enabled,
                    "MERGE_MODE": g_mode,
                    "KEY_VARS": "",
                    "RESULT_DATASET": out_name,
                    "NOTE": "multiple groups but no common keys, output separately"
                })
            continue

        final_df = reduce(
            lambda left, right: pd.merge(
                left,
                right,
                on=common_keys,
                how="outer",
                suffixes=("", "_DUP")
            ),
            mergeable_dfs
        )

        dup_cols = [c for c in final_df.columns if c.endswith("_DUP")]
        if dup_cols:
            final_df = final_df.drop(columns=dup_cols)

        # 保險補欄位
        if "STAT_DATASET" not in final_df.columns:
            final_df["STAT_DATASET"] = result_dataset

        merged_dataset_map[result_dataset] = final_df

        all_subitems = []
        group_modes = []

        for _, _, g_mode, g_names in group_results:
            group_modes.append(g_mode)
            all_subitems.extend(g_names)

        merge_logs.append({
            "FORM": form,
            "RAW_DATASETS": ", ".join(all_subitems),
            "STAT_DATASET": result_dataset,
            "SUBITEMS": ", ".join(all_subitems),
            "MERGE_ENABLED": enabled,
            "MERGE_MODE": "multi-stage",
            "KEY_VARS": ", ".join(common_keys),
            "RESULT_DATASET": result_dataset,
            "NOTE": f"group modes = {', '.join(group_modes)}; final = horizontal across groups"
        })

    # =========================================================
    # 保留沒被任何 FORM consume 到的 raw dataset
    # =========================================================
    for ds, df in raw_dataset_map.items():

        raw_ds = str(ds).strip().upper()

        if raw_ds not in all_used_subitems and raw_ds not in merged_dataset_map:

            tmp = _normalize_cols(df)
            tmp["RAW_DATASET"] = raw_ds
            tmp["STAT_DATASET"] = raw_ds

            merged_dataset_map[raw_ds] = tmp

    merge_log_df = pd.DataFrame(merge_logs, columns=[
        "FORM",
        "RAW_DATASETS",
        "STAT_DATASET",
        "SUBITEMS",
        "MERGE_ENABLED",
        "MERGE_MODE",
        "KEY_VARS",
        "RESULT_DATASET",
        "NOTE"
    ])

    return merged_dataset_map, merge_log_df



def get_source_datasets_for_merged(result_dataset, merge_log_df):

    if merge_log_df is None or merge_log_df.empty:
        return [result_dataset]

    result_dataset = str(result_dataset).strip().upper()

    tmp = merge_log_df[
        merge_log_df["RESULT_DATASET"].astype(str).str.upper() == result_dataset
    ]

    if tmp.empty:
        return [result_dataset]

    row = tmp.iloc[0]

    if "RAW_DATASETS" in row and pd.notna(row["RAW_DATASETS"]):
        txt = str(row["RAW_DATASETS"])
    else:
        txt = str(row.get("SUBITEMS", ""))

    vals = [x.strip().upper() for x in txt.split(",") if x.strip()]

    return vals if vals else [result_dataset]




def detect_date_variables(
    dataset_name,
    df,
    schema_df=None,
    raw_df=None,
    merge_log_df=None
):
    """
    判斷日期欄位（優先順序）：
    1. schema_df 的 DATATYPE
    2. raw_df 的 RAW_DATATYPE_STD
    3. fallback 欄名 heuristic
    """

    date_cols = set()
    dataset_name = str(dataset_name).strip().upper()

    # =================================================
    # 1️⃣ schema_df
    # =================================================
    if schema_df is not None and not schema_df.empty:

        tmp_schema = schema_df.copy()

        if "STAT_DATASET" in tmp_schema.columns:
            tmp_schema = tmp_schema[
                tmp_schema["STAT_DATASET"].astype(str).str.upper() == dataset_name
            ].copy()
        elif "RAW_DATASET" in tmp_schema.columns:
            tmp_schema = tmp_schema[
                tmp_schema["RAW_DATASET"].astype(str).str.upper() == dataset_name
            ].copy()

        for _, row in tmp_schema.iterrows():

            var = str(row.get("VARIABLE", "")).strip().upper()
            dtype = str(row.get("SCHEMA_DATATYPE_STD", "")).strip().upper()

            if var in df.columns:
                if dtype in ["DATE", "DATETIME", "TIME"]:
                    date_cols.add(var)

    # =================================================
    # 2️⃣ raw_df
    # =================================================
    if raw_df is not None and not raw_df.empty:

        source_datasets = get_source_datasets_for_merged(dataset_name, merge_log_df)

        for col in df.columns:
            cu = str(col).strip().upper()

            if cu in date_cols:
                continue

            tmp = raw_df[
                raw_df["VARIABLE"].astype(str).str.upper() == cu
            ].copy()

            if "RAW_DATASET" in tmp.columns:
                tmp = tmp[
                    tmp["RAW_DATASET"].astype(str).str.upper().isin(source_datasets)
                ]

            if not tmp.empty and "RAW_DATATYPE_STD" in tmp.columns:
                stds = tmp["RAW_DATATYPE_STD"].astype(str).str.upper().unique().tolist()

                if any(x in ["DATE", "DATETIME", "TIME"] for x in stds):
                    date_cols.add(cu)

    # =================================================
    # 3️⃣ fallback
    # =================================================
    for col in df.columns:

        cu = str(col).strip().upper()

        if cu in date_cols:
            continue

        if (
            cu.endswith("DAT")
            or cu.endswith("DTC")
            or cu.endswith("DATE")
            or cu.endswith("DT")
            or "STDAT" in cu
            or "ENDAT" in cu
        ):
            date_cols.add(cu)

    return sorted(date_cols)






SAS_DATE_EPOCH = date(1960, 1, 1)

def normalize_partial_date(value):
    """
    回傳:
      (date_char, sas_date_num)

    支援:
      YYYY          -> YYYY-01-01
      YYYY-MM       -> YYYY-MM-01
      YYYY-MM-DD    -> YYYY-MM-DD
      YYYY/MM       -> YYYY-MM-01
      YYYY/MM/DD    -> YYYY-MM-DD

    重要：
      判斷順序必須先 YYYY-MM-DD，再 YYYY-MM。
    """

    if pd.isna(value):
        return "", None

    s = str(value).strip()

    if s == "":
        return "", None

    s = s.replace("/", "-")

    parts = [p.strip() for p in re.split(r"[-\s]+", s) if p.strip() != ""]
    num_parts = [p for p in parts if p.isdigit()]

    if not num_parts:
        return "", None

    # -------------------------------------------------
    # YYYY-MM-DD：一定要先判斷
    # -------------------------------------------------
    if (
        len(num_parts) >= 3
        and len(num_parts[0]) == 4
        and len(num_parts[1]) in [1, 2]
        and len(num_parts[2]) in [1, 2]
    ):
        y = int(num_parts[0])
        m = int(num_parts[1])
        d = int(num_parts[2])
        s2 = f"{y:04d}-{m:02d}-{d:02d}"

    # -------------------------------------------------
    # YYYY-MM
    # -------------------------------------------------
    elif (
        len(num_parts) >= 2
        and len(num_parts[0]) == 4
        and len(num_parts[1]) in [1, 2]
    ):
        y = int(num_parts[0])
        m = int(num_parts[1])
        s2 = f"{y:04d}-{m:02d}-01"

    # -------------------------------------------------
    # YYYY
    # -------------------------------------------------
    elif len(num_parts) == 1 and len(num_parts[0]) == 4:
        y = int(num_parts[0])
        s2 = f"{y:04d}-01-01"

    else:
        return "", None

    try:
        dt = datetime.strptime(s2, "%Y-%m-%d").date()
    except Exception:
        return "", None

    sas_num = (dt - SAS_DATE_EPOCH).days

    return s2, int(sas_num)



def normalize_time_value(value):
    """
    回傳:
      (time_char, sas_time_num)

    支援:
      1030       -> 10:30
      103000     -> 10:30:00
      10:30      -> 10:30
      10:30:15   -> 10:30:15

    其他不合法 -> ("", None)
    """

    if pd.isna(value):
        return "", None

    s = str(value).strip()
    if s == "":
        return "", None

    # 只保留數字與冒號
    s = re.sub(r"[^0-9:]", "", s)

    # =========================
    # Hour only
    # =========================
    if re.fullmatch(r"\d{1,2}", s):
        hh = int(s)
        if hh > 23:
            return "", None

        time_c = f"{hh:02d}"         # _C 不補
        sas_time_num = hh * 3600     # _N 補成 HH:00:00
        return time_c, sas_time_num

    # =========================
    # 10:
    # =========================
    if re.fullmatch(r"\d{1,2}:", s):
        hh = int(s[:-1])
        if hh > 23:
            return "", None

        time_c = f"{hh:02d}"         # _C 不補
        sas_time_num = hh * 3600
        return time_c, sas_time_num

    # =========================
    # HHMM
    # =========================
    if re.fullmatch(r"\d{4}", s):
        hh = int(s[:2])
        mm = int(s[2:])
        if hh > 23 or mm > 59:
            return "", None

        time_c = f"{hh:02d}:{mm:02d}"
        sas_time_num = hh * 3600 + mm * 60
        return time_c, sas_time_num

    # =========================
    # HHMMSS
    # =========================
    if re.fullmatch(r"\d{6}", s):
        hh = int(s[:2])
        mm = int(s[2:4])
        ss = int(s[4:])
        if hh > 23 or mm > 59 or ss > 59:
            return "", None

        time_c = f"{hh:02d}:{mm:02d}:{ss:02d}"
        sas_time_num = hh * 3600 + mm * 60 + ss
        return time_c, sas_time_num

    # =========================
    # HH:MM
    # =========================
    if re.fullmatch(r"\d{1,2}:\d{1,2}", s):
        hh, mm = s.split(":")
        hh = int(hh)
        mm = int(mm)
        if hh > 23 or mm > 59:
            return "", None

        time_c = f"{hh:02d}:{mm:02d}"
        sas_time_num = hh * 3600 + mm * 60
        return time_c, sas_time_num

    # =========================
    # HH:MM:SS
    # =========================
    if re.fullmatch(r"\d{1,2}:\d{1,2}:\d{1,2}", s):
        hh, mm, ss = s.split(":")
        hh = int(hh)
        mm = int(mm)
        ss = int(ss)
        if hh > 23 or mm > 59 or ss > 59:
            return "", None

        time_c = f"{hh:02d}:{mm:02d}:{ss:02d}"
        sas_time_num = hh * 3600 + mm * 60 + ss
        return time_c, sas_time_num

    return "", None



SAS_DATETIME_EPOCH = datetime(1960, 1, 1, 0, 0, 0)

def combine_iso_datetime(date_char, time_char):
    """
    回傳:
      (dtc_char, sas_datetime_num)

    date_char: YYYY-MM-DD
    time_char: HH:MM or HH:MM:SS
    """

    d = str(date_char).strip() if pd.notna(date_char) else ""
    t = str(time_char).strip() if pd.notna(time_char) else ""

    if d == "":
        return "", None

    # 沒 time，就只回 date char；numeric datetime 也不產
    if t == "":
        return d, None
    
    dtc_c = f"{d}T{t}"

    # _N：補齊 date/time
    # date
    d_parts = d.split("-")
    if len(d_parts) == 1:          # YYYY
        d_for_num = f"{d}-01-01"
    elif len(d_parts) == 2:        # YYYY-MM
        d_for_num = f"{d}-01"
    else:
        d_for_num = d

    # time
    if re.fullmatch(r"\d{2}", t):             # HH
        t_for_num = f"{t}:00:00"
    elif re.fullmatch(r"\d{2}:\d{2}", t):     # HH:MM
        t_for_num = f"{t}:00"
    else:
        t_for_num = t

    try:
        dt_val = datetime.strptime(f"{d_for_num}T{t_for_num}", "%Y-%m-%dT%H:%M:%S")
    except Exception:
        return dtc_c, None

    sas_dt_num = int((dt_val - SAS_DATETIME_EPOCH).total_seconds())
    return dtc_c, sas_dt_num





def add_date_derivatives(dataset_name, df, schema_df=None, raw_df=None, merge_log_df=None) -> pd.DataFrame:

    out = df.copy()

    # =================================================
    # 1) DATE / Partial Date -> _C / _N
    # =================================================
    date_cols = detect_date_variables(
        dataset_name=dataset_name,
        df=out,
        schema_df=schema_df,
        raw_df=raw_df,
        merge_log_df=merge_log_df
    )

    for col in date_cols:
        c_col = f"{col}_C"
        n_col = f"{col}_N"

        chars = []
        nums = []

        for val in out[col]:
            cval, nval = normalize_partial_date(val)
            chars.append(cval)
            nums.append(nval)

        out[c_col] = chars
        out[n_col] = nums

    # =================================================
    # 2) TIME -> _C / _N
    #    依 schema DATATYPE = TIME
    # =================================================
    time_cols = []

    if schema_df is not None and not schema_df.empty:

        tmp_schema = schema_df.copy()

        if "STAT_DATASET" in tmp_schema.columns:
            tmp_schema = tmp_schema[
                tmp_schema["STAT_DATASET"].astype(str).str.upper()
                == str(dataset_name).strip().upper()
            ].copy()
        elif "RAW_DATASET" in tmp_schema.columns:
            tmp_schema = tmp_schema[
                tmp_schema["RAW_DATASET"].astype(str).str.upper()
                == str(dataset_name).strip().upper()
            ].copy()


        if "VARIABLE" in tmp_schema.columns and "SCHEMA_DATATYPE_STD" in tmp_schema.columns:
            for _, row in tmp_schema.iterrows():
                var = str(row.get("VARIABLE", "")).strip().upper()
                dtype = str(row.get("SCHEMA_DATATYPE_STD", "")).strip().upper()

                if dtype == "TIME":
                    for c in out.columns:
                        if str(c).strip().upper() == var:
                            time_cols.append(c)
                            break

    if not time_cols:
        for c in out.columns:
            if str(c).strip().upper().endswith("TIM"):
                time_cols.append(c)

    time_cols = sorted(set(time_cols))


    for col in time_cols:
        c_col = f"{col}_C"
        n_col = f"{col}_N"

        chars = []
        nums = []

        for val in out[col]:
            cval, nval = normalize_time_value(val)
            chars.append(cval)
            nums.append(nval)

        out[c_col] = chars
        out[n_col] = nums

    # =================================================
    # 3) DAT + TIM -> DTC_C / DTC_N
    #    例如 AESTDAT + AESTTIM -> AESTDTC_C / AESTDTC_N
    # =================================================
    raw_cols_snapshot = list(out.columns)

    for col in raw_cols_snapshot:
        col_upper = str(col).strip().upper()

        # 只處理原始 DAT 欄位，不吃衍生欄
        if not col_upper.endswith("DAT"):
            continue

        if col_upper.endswith("_C") or col_upper.endswith("_N"):
            continue


        prefix = col[:-3]          # AESTDAT -> AEST
        time_col = f"{prefix}TIM"  # AESTTIM

        date_c_col = f"{col}_C"        # AESTDAT_C
        time_c_col = f"{time_col}_C"   # AESTTIM_C

        dtc_c_col = f"{prefix}DTC_C"   # AESTDTC_C
        dtc_n_col = f"{prefix}DTC_N"   # AESTDTC_N


        if date_c_col not in out.columns:
            continue

        dtc_chars = []
        dtc_nums = []

        for i in range(len(out)):
            date_char = out.iloc[i][date_c_col]

            if time_c_col in out.columns:
                time_char = out.iloc[i][time_c_col]
            else:
                time_char = None

            dtc_char, dtc_num = combine_iso_datetime(date_char, time_char)

            dtc_chars.append(dtc_char)
            dtc_nums.append(dtc_num)

        out[dtc_c_col] = dtc_chars
        out[dtc_n_col] = dtc_nums

    
    # =================================================
    # 4) 欄位順序：讓衍生欄位跟在原欄位後
    # =================================================

    original_cols = list(df.columns)

    derived_cols = []

    for col in original_cols:

        col_upper = str(col).strip().upper()

        # -------------------------
        # DATE
        # -------------------------
        c_col = f"{col}_C"
        n_col = f"{col}_N"

        if c_col in out.columns:
            derived_cols.append(c_col)

        if n_col in out.columns:
            derived_cols.append(n_col)

        # -------------------------
        # TIME
        # -------------------------
        if col_upper.endswith("TIM"):

            tim_c = f"{col}_C"
            tim_n = f"{col}_N"

            if tim_c in out.columns:
                derived_cols.append(tim_c)

            if tim_n in out.columns:
                derived_cols.append(tim_n)

        # -------------------------
        # DAT → DTC
        # -------------------------
        if col_upper.endswith("DAT"):

            prefix = col[:-3]

            dtc_c = f"{prefix}DTC_C"
            dtc_n = f"{prefix}DTC_N"

            if dtc_c in out.columns:
                derived_cols.append(dtc_c)

            if dtc_n in out.columns:
                derived_cols.append(dtc_n)


    # 去重（避免重複）
    derived_cols_unique = []
    seen = set()

    for c in derived_cols:
        if c not in seen:
            derived_cols_unique.append(c)
            seen.add(c)


    # 原始欄位（排除已是 derived）
    base_cols = [c for c in out.columns if c not in derived_cols_unique]

    # 最終順序
    out = out[base_cols + derived_cols_unique]


    return out



def enrich_merged_dataset_map(
    merged_dataset_map: dict,
    visit_map_df: pd.DataFrame | None = None,
    schema_df: pd.DataFrame | None = None,
    raw_df: pd.DataFrame | None = None,
    merge_log_df: pd.DataFrame | None = None
):

    """
    merge 後的整理：
    1. 依 Folder sheet 補 / 標準化 VISIT, VISIT_DVP, VISITNUM
    2. 日期 / 時間 / datetime 派生 (_C / _N / DTC)
    """

    out = {}

    for name, df in merged_dataset_map.items():
        if df is None:
            out[name] = df
            continue

        df2 = df.copy()

        # 1) visit mapping
        if visit_map_df is not None and not visit_map_df.empty and "VISIT" in df2.columns:

            visit_lookup = visit_map_df.copy()

            # normalize
            visit_lookup["VISIT"] = visit_lookup["VISIT"].astype(str).str.strip()
            df2["VISIT"] = df2["VISIT"].astype(str).str.strip()

            # merge
            keep_cols = ["VISIT", "VISIT_DVP"]
            if "VISITNUM" in visit_lookup.columns:
                keep_cols.append("VISITNUM")

            visit_lookup = visit_lookup[keep_cols].drop_duplicates()

            df2 = df2.merge(visit_lookup, on="VISIT", how="left")

        else:
            # fallback（萬一沒有 visit_map_df）
            if "VISIT_DVP" not in df2.columns:
                df2["VISIT_DVP"] = ""
            if "VISITNUM" not in df2.columns:
                df2["VISITNUM"] = pd.N


        # 2) date derivatives
        df2 = add_date_derivatives(
            dataset_name=name,
            df=df2,
            schema_df=schema_df,
            raw_df=raw_df,
            merge_log_df=merge_log_df
        )


        # 欄位順序調整（VISIT_DVP 和 VISITNUM 放原本 VISIT 後）
        cols = list(df2.columns)

        prefix_cols = []
        if "SOURCEDN" in cols:
            prefix_cols.append("SOURCEDN")

        
        for c in ["VISIT", "VISIT_DVP", "VISITNUM"]:
            if c in cols:
                prefix_cols.append(c)

        prefix_set = set(prefix_cols)


       # -------------------------
        # ② 原始欄位（排除 derived）
        # -------------------------
        original_cols = []
        for c in cols:
            cu = str(c).upper()

            if c in prefix_set:
                continue

            # 排除所有 derived
            if (
                cu.endswith("_C")
                or cu.endswith("_N")
                or "DTC_" in cu
            ):
                continue

            original_cols.append(c)

        # -------------------------
        # ③ derived（依 DAT/TIM 順序）
        # -------------------------
        derived_cols = []

        for c in cols:
            cu = str(c).upper()

            # ---------- DATE ----------
            if cu.endswith("DAT"):

                if f"{c}_C" in df2.columns:
                    derived_cols.append(f"{c}_C")

                if f"{c}_N" in df2.columns:
                    derived_cols.append(f"{c}_N")

                prefix = c[:-3]

                dtc_c = f"{prefix}DTC_C"
                dtc_n = f"{prefix}DTC_N"

                if dtc_c in df2.columns:
                    derived_cols.append(dtc_c)

                if dtc_n in df2.columns:
                    derived_cols.append(dtc_n)

            # ---------- TIME ----------
            if cu.endswith("TIM"):

                if f"{c}_C" in df2.columns:
                    derived_cols.append(f"{c}_C")

                if f"{c}_N" in df2.columns:
                    derived_cols.append(f"{c}_N")

        # 去重
        seen = set()
        derived_cols_unique = []

        for c in derived_cols:
            if c not in seen:
                derived_cols_unique.append(c)
                seen.add(c)

        # -------------------------
        # ✅ 最終欄位順序
        # -------------------------
        final_cols = prefix_cols + original_cols + derived_cols_unique
        drop_cols = ["RAW_DATASET", "STAT_DATASET"]
        final_cols = [c for c in final_cols if c not in drop_cols]

        df2 = df2[final_cols]

        out[name] = df2

    return out






# =========================================================
# 主流程
# =========================================================

def run_pipeline(raw_dataset_map, raw_spec, system_vars=None, merge_config=None):
    """
    回傳：
    - schema_df
    - raw_df
    - compare_df
    - raw_preview_map
    - raw_dataset_map
    - merged_dataset_map
    - merge_log_df
    """

    raw_df, raw_preview_map, raw_dataset_map = build_raw_metadata_and_preview(raw_dataset_map)

    schema_df = raw_spec["variable"].copy()

    # normalize
    schema_df["FORM"] = schema_df["FORM"].astype(str).str.strip().str.upper()
    schema_df["RAW_DATASET"] = schema_df["RAW_DATASET"].astype(str).str.strip().str.upper()
    schema_df["STAT_DATASET"] = schema_df["STAT_DATASET"].astype(str).str.strip().str.upper()
    schema_df["VARIABLE"] = schema_df["VARIABLE"].astype(str).str.strip().str.upper()

    compare_df = compare_schema_raw(
        schema_df,
        raw_df,
        system_vars=system_vars
    )

    if "VARIABLE_ORDER" in compare_df.columns:
        compare_df["VARIABLE_ORDER"] = (
            pd.to_numeric(compare_df["VARIABLE_ORDER"], errors="coerce")
            .astype("Int64")
        )

        compare_df = compare_df.sort_values(
            by=["FORM", "RAW_DATASET", "VARIABLE_ORDER", "VARIABLE"],
            na_position="last"
        ).reset_index(drop=True)

    if merge_config is None:
        merge_config = {
            "enabled": True,
            "default_key_vars": system_vars,
            "form_rules": {}
        }

    merged_dataset_map, merge_log_df = merge_subitems(
        schema_df=schema_df,
        raw_dataset_map=raw_dataset_map,
        merge_config=merge_config
    )

    return (
        schema_df,
        raw_df,
        compare_df,
        raw_preview_map,
        raw_dataset_map,
        merged_dataset_map,
        merge_log_df
    )





def process_raw_data(schema_file, zip_file, system_vars):

    raw_spec = build_rawdata_spec(schema_file)

    raw_dataset_map = load_sas_from_zip(zip_file)

    (
        schema_df,
        raw_df,
        compare_df,
        raw_preview_map,
        raw_dataset_map,
        merged_dataset_map,
        merge_log_df
    ) = run_pipeline(
        raw_dataset_map=raw_dataset_map,
        raw_spec=raw_spec,
        system_vars=system_vars
    )

    visit_map_df = build_visit_mapping_table(
        raw_spec["visit"],
        raw_dataset_map
    )

    merged_dataset_map = enrich_merged_dataset_map(
        merged_dataset_map,
        visit_map_df,
        schema_df=schema_df,
        raw_df=raw_df,
        merge_log_df=merge_log_df
    )


    return {
        "raw_spec": raw_spec,
        "schema_df": schema_df,
        "raw_df": raw_df,
        "compare_df": compare_df,
        "raw_preview_map": raw_preview_map,
        "raw_dataset_map": raw_dataset_map,
        "merged_dataset_map": merged_dataset_map,
        "merge_log_df": merge_log_df,
        "visit_map_df": visit_map_df
    }




