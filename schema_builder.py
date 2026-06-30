import pandas as pd
import re



# =========================================================
# Helper
# =========================================================
def normalize_colname(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().lower()

    s = re.sub(r"\[.*?\]", "", s)
    s = re.sub(r"\s+", " ", s)

    return s.strip()


def find_col(df, keywords):
    for c in df.columns:

        clean = normalize_colname(c)

        for kw in keywords:
            if kw in clean:
                return c
    return None



# Datatype normalize（給 Data Check 用）
def normalize_dtype(x):

    if pd.isna(x):
        return "TEXT"

    s = str(x).lower()

    if "datetime" in s:
        return "DATETIME"
    elif "date" in s:
        return "DATE"
    elif "time" in s:
        return "TIME"
    elif any(k in s for k in ["int", "num", "float", "decimal"]):
        return "NUMERIC"
    else:
        return "TEXT"



# =========================================================
# 1) detect_header_row
# =========================================================
def detect_header_row(df, keywords, max_rows=40):
    """
    keywords: list[str]
    只要某一列同時包含所有 keywords，就視為 header row
    """

    scan_n = min(len(df), max_rows)
    keywords = [normalize_colname(k) for k in keywords]

    for i in range(scan_n):

        row_values = [
            normalize_colname(x)
            for x in df.iloc[i]
        ]

        # 每個 keyword 至少要在該列某個 cell 中出現
        if all(any(k in cell for cell in row_values) for k in keywords):
            return i

    return None


# =========================================================
# 2) read_sheet_with_detected_header
# =========================================================
def read_sheet_with_detected_header(xl, sheet_name, keywords):
    """
    先 preview(header=None) → detect header row → 再重讀
    keywords 由外部傳入，不在這裡寫條件
    """

    preview = pd.read_excel(
        xl,
        sheet_name=sheet_name,
        header=None,
        dtype=object
    )

    header_row = detect_header_row(preview, keywords=keywords)

    if header_row is None:
        return None

    df = pd.read_excel(
        xl,
        sheet_name=sheet_name,
        header=header_row,
        dtype=object
    )

    df.columns = [str(c).strip() for c in df.columns]

    return df


# =========================================================
# 3) parse_domain_sheet
# =========================================================
def parse_domain_sheet(df, sheet_name):
    """
    解析 AE / DM / VS ... 等 CRF domain sheet

    來源欄位可能包含：
    - Field OID
    - Sub-item
    - Field Name
    - Field Type
    - Option_Displayed Value
    """

    if df is None or df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    field_oid_col   = find_col(df, ["field oid", "item oid", "variable"])
    subitem_col     = find_col(df, ["sub-item"])
    sn_col          = find_col(df, ["s/n", "sn"])
    field_name_col  = find_col(df, ["field name", "item name", "label"])
    field_type_col  = find_col(df, ["field type", "fieldtype"])
    option_col      = find_col(df, ["option", "display"])

    if field_oid_col is None:
        return pd.DataFrame(), pd.DataFrame()
    
    
    # 若沒有 S/N，fallback 用 row order
    if sn_col is None:
        df["_AUTO_SN"] = range(1, len(df) + 1)
        sn_col = "_AUTO_SN"

    variable_rows = []
    code_rows = []

    for _, r in df.iterrows():

        sn_val = r.get(sn_col) if sn_col else None
        var = r.get(field_oid_col)
        
        if pd.isna(var) or str(var).strip() == "":
            continue
        
        var = str(var).strip().upper()

        # DATASET 優先取 Sub-item，若無則用 sheet name
        raw_dataset = (
            str(r.get(subitem_col)).strip().upper()
            if subitem_col and pd.notna(r.get(subitem_col)) and str(r.get(subitem_col)).strip() != ""
            else str(sheet_name).strip().upper()
        )

        label = r.get(field_name_col)
        field_type = r.get(field_type_col)
        label_str = str(label).strip() if pd.notna(label) else ""


        # -------------------------
        # Variable List
        # -------------------------
        variable_rows.append({
            "FORM": str(sheet_name).strip(),
            "RAW_DATASET": raw_dataset,
            "STAT_DATASET": "",
            "VARIABLE_ORDER": sn_val,
            "VARIABLE": var,
            "LABEL": str(label).strip() if pd.notna(label) else "",
            "SCHEMA_DATATYPE": str(field_type).strip() if pd.notna(field_type) else "",
            "SCHEMA_DATATYPE_STD": normalize_dtype(field_type),
            "VALUELIST": ""
        })

        # -------------------------
        # Value List
        # -------------------------
        codes_to_add = []

        options = r.get(option_col)

        # 1) Option values
        if pd.notna(options) and str(options).strip():

            opts = re.split(r"[\n;,]", str(options))

            for opt in opts:
                opt = opt.strip()

                if not opt:
                    continue

                opt_upper = opt.upper()

                # filter placeholder
                if opt_upper == "XXX":
                    continue

                if "XXX" in opt_upper:
                    continue

                codes_to_add.append(opt)


        # 2) Constant fallback（只有沒 option 時才看 label）
        if pd.notna(field_type):

            ft = str(field_type).lower()

            if "constant" in ft and label_str:

                # 只有在 option 沒有帶出任何 code 時，才 fallback 用 label
                if not codes_to_add:

                    label_upper = label_str.upper()

                    # template label 直接跳過
                    if "XXX" not in label_upper:

                        code_clean = re.sub(r"\(.*?\)", "", label_str)
                        val = code_clean.strip()

                        if val:
                            codes_to_add.append(val)

        # 3) 統一 append
        for seq, val in enumerate(codes_to_add, start=1):
            code_rows.append({
                "FORM": str(sheet_name).strip(),
                "RAW_DATASET": raw_dataset,
                "VARIABLE_ORDER": sn_val,
                "VARIABLE": var,
                "VALUELIST": var,
                "VALUE_ORDER": seq,
                "VALUE": val
            })


    variable_df = pd.DataFrame(variable_rows)
    code_df = pd.DataFrame(code_rows).drop_duplicates()
    
    if not variable_df.empty:

        variable_df = variable_df.sort_values(
            by=["FORM", "VARIABLE_ORDER"],
            na_position="last"
        )

        # optional: reset index
        variable_df = variable_df.reset_index(drop=True)

    if not code_df.empty:

        # 轉數值（排序用）
        code_df["VARIABLE_ORDER_NUM"] = pd.to_numeric(code_df["VARIABLE_ORDER"], errors="coerce")
        code_df["VALUE_ORDER_NUM"] = pd.to_numeric(code_df["VALUE_ORDER"], errors="coerce")

        # Step 1：先按原順序排序（確保 earliest 在前）
        code_df = code_df.sort_values(
            by=["FORM", "RAW_DATASET", "VALUELIST", "VARIABLE_ORDER_NUM", "VALUE_ORDER_NUM"],
            na_position="last"
        )

        # Step 2：去重（保留 first）
        code_df = code_df.drop_duplicates(subset=["VALUELIST", "VALUE"],keep="first")

        # Step 3：重新編 VALUE_ORDER
        code_df["VALUE_ORDER"] = (code_df.groupby("VALUELIST").cumcount() + 1)
        code_df = code_df.drop(columns=["VARIABLE_ORDER_NUM", "VALUE_ORDER_NUM"], errors="ignore")

        # 最後整理排序
        code_df = code_df[[
            "VALUELIST",
            "VALUE_ORDER",
            "VALUE"
        ]]

        code_df = code_df.sort_values(
            by=["VALUELIST", "VALUE_ORDER"],
        ).reset_index(drop=True)

    return variable_df, code_df


# =========================================================
# 4) parse_folder_sheet
# =========================================================
def parse_folder_sheet(df):
    """
    Visit List 只從 Folder sheet 來
    """

    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    abbr_col     = find_col(df, ["abbrev", "abbreviation"])
    full_col     = find_col(df, ["full", "full term"])
    repeat_col   = find_col(df, ["repeat"])

    if abbr_col is None or full_col is None:
        return pd.DataFrame()

    visit_rows = []

    # 直接用 row order 當 VISIT_ORDER
    for idx, r in df.iterrows():

        sn_val = idx + 1

        visit = str(r.get(full_col)).strip() if pd.notna(r.get(full_col)) else ""
        abbr  = str(r.get(abbr_col)).strip() if pd.notna(r.get(abbr_col)) else ""

        if not visit:
            continue

        visit_rows.append({
            "VISIT_ORDER": sn_val,
            "VISIT": visit,
            "VISIT_DVP": abbr,
            "REPEAT_FOLDER": str(r.get(repeat_col)).strip() if repeat_col and pd.notna(r.get(repeat_col)) else ""
        })

    visit_df = pd.DataFrame(visit_rows)

    if not visit_df.empty:
        visit_df = visit_df.sort_values("VISIT_ORDER").reset_index(drop=True)

    return visit_df





# =========================================================
# 5) parse_soa_sheet
# =========================================================
def parse_soa_sheet(df, visit_df, variable_df=None, sheet_name="SoA"):
    """
    用 SoA 產出各 CRF domain 對應的 Visit

    輸出：
      FORM
      RAW_DATASET
      FORM_NAME
      VISIT
      VISIT_DVP
      VISIT_ORDER
    """

    if df is None or df.empty:
        return pd.DataFrame()

    if visit_df is None or visit_df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # -------------------------
    # 找 row metadata 欄位
    # -------------------------
    abbr_col = find_col(df, ["abbreviation"])
    crf_name_col = find_col(df, ["crf name"])

    if abbr_col is None:
        return pd.DataFrame(columns=[
            "FORM", "FORM_NAME", "RAW_DATASET", "STAT_DATASET", "VISIT", "VISIT_DVP", "VISIT_ORDER"
        ])

    # -------------------------
    # visit_df 準備：用 Folder 的 VISIT_DVP 當 SoA 欄名對照
    # -------------------------
    visit_df2 = visit_df.copy()

    # 如果你 visit_df 目前只有 _SN，這裡轉成正式順序欄
    if "VISIT_ORDER" not in visit_df2.columns:
        if "_SN" in visit_df2.columns:
            visit_df2["VISIT_ORDER"] = visit_df2["_SN"]
        else:
            visit_df2["VISIT_ORDER"] = range(1, len(visit_df2) + 1)

    visit_df2["VISIT_DVP"] = visit_df2["VISIT_DVP"].astype(str).str.strip().str.upper()
    visit_df2["VISIT"] = visit_df2["VISIT"].astype(str).str.strip()

    # lookup：SoA 欄名 -> Folder visit meta
    visit_lookup = {}
    for _, row in visit_df2.iterrows():
        k = str(row["VISIT_DVP"]).strip().upper()
        if k:
            visit_lookup[k] = {
                "VISIT": row["VISIT"],
                "VISIT_DVP": row["VISIT_DVP"],
                "VISIT_ORDER": row["VISIT_ORDER"]
            }

    # -------------------------
    # 找 SoA 真正的 visit 欄（用 VISIT_DVP 對）
    # -------------------------
    soa_visit_cols = []
    for c in df.columns:
        c_norm = str(c).strip().upper()
        if c_norm in visit_lookup:
            soa_visit_cols.append(c)

    if not soa_visit_cols:
        return pd.DataFrame(columns=[
            "FORM", "FORM_NAME", "RAW_DATASET", "STAT_DATASET", "VISIT", "VISIT_DVP", "VISIT_ORDER"
        ])

    # -------------------------
    # 判斷 cell 是否代表「有收資料」
    # -------------------------
    def is_collected(x):
        if pd.isna(x):
            return False

        s = str(x).strip()
        if s == "":
            return False

        s_up = s.upper()

        # 常見表記
        checked_values = {
            "X", "Y", "YES", "TRUE", "1",
            "✓", "✔", "■", "●", "☑", "☒", "▣"
        }

        unchecked_values = {
            "□", "☐", "0", "FALSE", "NO", "N"
        }

        if s_up in checked_values:
            return True
        if s_up in unchecked_values:
            return False

        # 不是空白、不是明確未勾選，就先視為 collected
        return True

    # -------------------------
    # 展開成 domain x visit
    # -------------------------
    out_rows = []

    for _, r in df.iterrows():

        form = r.get(abbr_col)
        if pd.isna(form) or str(form).strip() == "":
            continue

        form = str(form).strip().upper()
        crf_name = str(r.get(crf_name_col)).strip() if crf_name_col and pd.notna(r.get(crf_name_col)) else ""

        for visit_col in soa_visit_cols:
            cell = r.get(visit_col)

            if not is_collected(cell):
                continue

            meta = visit_lookup[str(visit_col).strip().upper()]

            out_rows.append({
                "FORM": form,
                "FORM_NAME": crf_name,
                "VISIT": meta["VISIT"],
                "VISIT_DVP": meta["VISIT_DVP"],
                "VISIT_ORDER": meta["VISIT_ORDER"]
            })

    domain_visit_df = pd.DataFrame(out_rows)

    if domain_visit_df.empty:
        return pd.DataFrame(columns=[
            "FORM", "FORM_NAME", "RAW_DATASET", "STAT_DATASET", "VISIT", "VISIT_DVP", "VISIT_ORDER"
        ])
    

    if variable_df is not None and not variable_df.empty:

        form_dataset_map = (
            variable_df[["FORM", "RAW_DATASET", "STAT_DATASET"]]
            .drop_duplicates()
        )

        domain_visit_df = domain_visit_df.merge(
            form_dataset_map,
            on="FORM",
            how="left"
        )
    else:       
        domain_visit_df["RAW_DATASET"] = None
        domain_visit_df["STAT_DATASET"] = None


    # =========================================================
    # 去重：同 FORM + VISIT 只保留 FORM==DATASET
    # =========================================================
    if not domain_visit_df.empty:

        # 先標記 match（FORM == DATASET）
        domain_visit_df["MATCH"] = (
            domain_visit_df["FORM"] == domain_visit_df["STAT_DATASET"]
        )

        # 排序：match優先
        domain_visit_df = domain_visit_df.sort_values(
            by=["FORM", "VISIT", "MATCH"],
            ascending=[True, True, False]
        )

        # group內只保留第一筆
        domain_visit_df = (
            domain_visit_df
            .drop_duplicates(subset=["FORM", "VISIT"], keep="first")
            .reset_index(drop=True)
        )

        # cleanup
        domain_visit_df = domain_visit_df.drop(columns=["MATCH"])
        
    domain_visit_df["VISIT_ORDER"] = pd.to_numeric(
        domain_visit_df["VISIT_ORDER"], errors="coerce"
    )

    domain_visit_df = domain_visit_df.sort_values(
        by=["STAT_DATASET", "VISIT_ORDER"],
        na_position="last"
    ).reset_index(drop=True)

    cols = [
        "STAT_DATASET",
        "FORM_NAME",
        "VISIT",
        "VISIT_DVP",
        "VISIT_ORDER"
    ]

    cols = [c for c in cols if c in domain_visit_df.columns]

    domain_visit_df = domain_visit_df[cols]

    return domain_visit_df




# =========================================================
# 5) assign_stat_dataset
# =========================================================
def assign_stat_dataset(variable_df):
    """
    依 FORM 判斷是否需要 merge dataset
    """

    if variable_df is None or variable_df.empty:
        return variable_df

    df = variable_df.copy()

    df["FORM"] = df["FORM"].str.upper().str.strip()
    df["RAW_DATASET"] = df["RAW_DATASET"].str.upper().str.strip()

    # 每個FORM有幾個dataset
    form_ds_count = (
        df.groupby("FORM")["RAW_DATASET"]
        .nunique()
        .to_dict()
    )

    df["STAT_DATASET"] = df.apply(
        lambda r: (
            r["FORM"] if form_ds_count.get(r["FORM"], 0) > 1
            else r["RAW_DATASET"]
        ),
        axis=1
    )

    return df



# =========================================================
# 6) build_rawdata_spec
# =========================================================
def build_rawdata_spec(schema_file):
    """
    eCRF Schema (multi-sheet) -> Raw Data SPEC
      - Variable List
      - Code List
      - Visit List
      - Domain Visit List

    規則：
    - Folder sheet → Visit List
    - SoA sheet → Domain Visit List
    - 其他 domain sheets → Variable / Code
    """

    xl = pd.ExcelFile(schema_file)

    variable_list = []
    code_list = []
    visit_df = pd.DataFrame(columns=["VISIT", "VISIT_DVP", "VISITNUM", "REPEAT_FOLDER"])
    domain_visit_list = []
    soa_sheets = []

    for sheet in xl.sheet_names:

        sheet_upper = str(sheet).strip().upper()

        # -------------------------
        # Folder -> Visit List
        # -------------------------
        if sheet_upper == "FOLDER":
            df = read_sheet_with_detected_header(
                xl,
                sheet,
                keywords=["abbrev", "full"]
            )

            if df is not None and not df.empty:
                visit_df = parse_folder_sheet(df)

            continue


        # -------------------------
        # SoA -> Domain-Visit mapping
        # -------------------------
        if "SOA" in sheet_upper or "SCHEDULE" in sheet_upper:
            df = read_sheet_with_detected_header(
                xl,
                sheet,
                keywords=["abbreviation", "crf name"]
            )
            
            if df is not None and not df.empty:
                soa_sheets.append((sheet, df))

            continue


        # -------------------------
        # Domain sheets -> Variable / Code
        # -------------------------
        df = read_sheet_with_detected_header(
            xl,
            sheet,
            keywords=["field oid", "field name"]
        )

        if df is None or df.empty:
            continue

        var_df, code_df = parse_domain_sheet(df, sheet)

        if not var_df.empty:
            variable_list.append(var_df)

        if not code_df.empty:
            code_list.append(code_df)

    variable_df = (
        pd.concat(variable_list, ignore_index=True)
        if variable_list
        else pd.DataFrame(columns=[
            "FORM", "RAW_DATASET", "STAT_DATASET", "VARIABLE_ORDER", "VARIABLE", "LABEL",
            "SCHEMA_DATATYPE", "SCHEMA_DATATYPE_STD", "VALUELIST"
        ])
    )

    variable_df = variable_df.sort_values(
        by=["FORM", "VARIABLE_ORDER"],
        na_position="last"
    )
    variable_df = assign_stat_dataset(variable_df)

    code_df = (
        pd.concat(code_list, ignore_index=True)
        if code_list
        else pd.DataFrame(columns=[
            "VALUELIST", "VALUE_ORDER", "VALUE"
        ])
    )

    code_df = code_df.sort_values(
        by=["VALUELIST", "VALUE_ORDER"],
    ).reset_index(drop=True)


    domain_visit_list = []

    for sheet, df in soa_sheets:

        if visit_df.empty:
            continue

        domain_visit_df = parse_soa_sheet(
            df,
            visit_df,
            variable_df=variable_df,
            sheet_name=sheet
        )

        if not domain_visit_df.empty:
            domain_visit_list.append(domain_visit_df)


    domain_visit_df = (
        pd.concat(domain_visit_list, ignore_index=True)
        if domain_visit_list
        else pd.DataFrame(columns=[
            "FORM", "FORM_NAME", "RAW_DATASET", "STAT_DATASET", "VISIT", "VISIT_DVP", "VISIT_ORDER"
        ])
    )

    
    return {
        "variable": variable_df,
        "code": code_df,
        "visit": visit_df,
        "domain_visit": domain_visit_df
    }
