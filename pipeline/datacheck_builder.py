import re
import copy
import pandas as pd
import numpy as np
from datetime import datetime

from schema_builder import (
    find_col,
    read_sheet_with_detected_header
)


# =========================================================
# DataCheck Builder v2
# Design:
#   Step 1: DVP -> dvp_review_df (review/execution spec table)
#   Step 2: per STAT_DATASET enrich refs once -> per rule/block evaluate
# =========================================================


# =========================================================
# Basic helpers
# =========================================================
def normalize_upper_cols(df):
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    out.columns = [str(c).strip().upper() for c in out.columns]
    return out


def normalize_text(x):
    if pd.isna(x):
        return ""
    s = str(x).replace("\n", " ").replace("\r", " ").replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


#統一 join key (SUBJID / VISIT)，確保合併時不會因為大小寫或空值失敗
def normalize_join_value_series(s):
    if s is None:
        return s
    out = s.astype(str).str.strip().str.upper()
    out = out.replace({
        "": pd.NA, " ": pd.NA, "NAN": pd.NA, "NONE": pd.NA,
        "<NA>": pd.NA, "NA": pd.NA, "N/A": pd.NA,
        "nan": pd.NA, "None": pd.NA, "none": pd.NA,
    })
    return out


def first_non_missing(series):
    if series is None:
        return pd.NA
    s = series.replace({
        "": pd.NA, " ": pd.NA, "None": pd.NA, "NONE": pd.NA,
        "none": pd.NA, "nan": pd.NA, "NaN": pd.NA, "NAN": pd.NA,
        "<NA>": pd.NA, "NA": pd.NA, "N/A": pd.NA,
    }).dropna()
    return pd.NA if len(s) == 0 else s.iloc[0]


def get_subject_key(df):
    if df is None or df.empty:
        return None
    cols = [str(c).strip().upper() for c in df.columns]
    for c in ["SUBJID", "USUBJID", "SUBJECT", "SUBJECTID"]:
        if c in cols:
            return c
    return None


def get_visit_key(df):
    if df is None or df.empty:
        return None
    cols = [str(c).strip().upper() for c in df.columns]
    for c in ["VISIT_DVP", "VISIT"]:
        if c in cols:
            return c
    return None


# 抓 schema 中所有合法變數，判斷 UNKNOWN_VAR
def get_valid_var_set(raw_spec):
    variable_df = raw_spec.get("variable") if isinstance(raw_spec, dict) else pd.DataFrame()
    if variable_df is None or variable_df.empty or "VARIABLE" not in variable_df.columns:
        return set()
    vals = variable_df["VARIABLE"].dropna().astype(str).str.strip().str.upper().tolist()
    # special runtime symbols/functions
    vals += ["X", "VISIT_DAY", "MAX", "MIN"]
    return set(vals)


# 變數對應的STAT_DATASET
def build_var_to_stat_map(raw_spec):
    variable_df = raw_spec.get("variable") if isinstance(raw_spec, dict) else pd.DataFrame()
    out = {}
    if variable_df is None or variable_df.empty:
        return out
    df = normalize_upper_cols(variable_df)
    if "VARIABLE" not in df.columns or "STAT_DATASET" not in df.columns:
        return out
    for _, r in df[["VARIABLE", "STAT_DATASET"]].dropna().drop_duplicates().iterrows():
        out.setdefault(str(r["VARIABLE"]).strip().upper(), set()).add(str(r["STAT_DATASET"]).strip().upper())
    return out

# STAT_DATASET的所有變數，判斷 target var
def build_stat_to_var_map(raw_spec):
    variable_df = raw_spec.get("variable") if isinstance(raw_spec, dict) else pd.DataFrame()
    out = {}
    if variable_df is None or variable_df.empty:
        return out
    df = normalize_upper_cols(variable_df)
    if "VARIABLE" not in df.columns or "STAT_DATASET" not in df.columns:
        return out
    for stat, g in df.groupby("STAT_DATASET"):
        out[str(stat).strip().upper()] = set(g["VARIABLE"].dropna().astype(str).str.strip().str.upper())
    return out








# =========================================================
# Load DVP
# =========================================================

# Helper: 標準化Form欄位
def normalize_form_name(s):
    if s is None:
        return ""

    s = str(s).upper().strip()

    # remove (xxx)
    s = re.sub(r"\(.*?\)", "", s)

    # remove plural
    if s.endswith("S"):
        s = s[:-1]

    return s.strip()


# Helper: 將 Excel 儲存格中的多行文字拆成 list
def split_multiline_cell(val):
    if pd.isna(val):
        return []

    s = str(val).strip()
    if s == "":
        return []

    parts = re.split(r"[\r\n]+", s)
    parts = [p.strip() for p in parts if p.strip() != ""]

    return parts if parts else [s]


# Main: 匯入/整理DVP (這步會先展開Form，因為要併DATASET)
def parse_dvp_file(file, visit_df=None):
    """
    讀 DVP 並統一欄位成:
      RULE
      VISIT
      DATASET
      FORM
      FIELD
      NOTE_
      PD
      CONDITION
      METHOD
    """

    xl = pd.ExcelFile(file, engine="openpyxl")

    all_rows = []

    for sheet in xl.sheet_names:
        
        sheet_upper = str(sheet).strip().upper()

        if "DVP" not in sheet_upper and "DATA VALIDATION PLAN" not in sheet_upper:
            continue

        df = read_sheet_with_detected_header(
            xl,
            sheet,
            keywords=["form", "condition", "message"]
        )

        if df is None or df.empty:
            continue

        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        
        rule_col = None
        for c in df.columns:
            cname = str(c).strip().lower()
            if cname in ["edit no.", "edit no"]:
                rule_col = c
                break

        visit_col = find_col(df, ["visit", "folder"])
        form_col = find_col(df, ["form"])
        field_col = find_col(df, ["field"])
        message_col = find_col(df, ["message"])
        pd_col = find_col(df, ["deviation"])
        condition_col = find_col(df, ["condition"])
        method_col = find_col(df, ["method"])

        out = pd.DataFrame({
            "RULE": df[rule_col] if rule_col else pd.NA,
            "VISIT": df[visit_col] if visit_col else pd.NA,
            "FORM": df[form_col] if form_col else pd.NA,
            "FIELD": df[field_col] if field_col else pd.NA,
            "NOTE": df[message_col] if message_col else pd.NA,
            "PD": df[pd_col] if pd_col else pd.NA,
            "CONDITION": df[condition_col] if condition_col else pd.NA,
            "METHOD": df[method_col] if method_col else pd.NA,
        })
        
        # Filters
        out = out[out["CONDITION"].notna()].copy()
        out["CONDITION"] = out["CONDITION"].astype(str).str.strip()
        out = out[out["CONDITION"] != ""]
        out = out[out["RULE"].notna() & (out["RULE"].astype(str).str.strip() != "")]
        out = out[out["METHOD"].astype(str).str.strip().str.upper().eq("SAS")]

        
        # =================================================
        # FORM 換行拆成多筆
        # =================================================
        out["FORM"] = out["FORM"].apply(split_multiline_cell)
        out = out.explode("FORM", ignore_index=True)

        out["FORM"] = out["FORM"].astype(str).str.strip()
        out = out[out["FORM"] != ""].copy()

        all_rows.append(out)

    if not all_rows:
        result = pd.DataFrame(columns=[
            "RULE", "VISIT", "FORM", "FIELD",
            "NOTE", "PD", "CONDITION", "METHOD"
        ])
    else:
        result = pd.concat(all_rows, ignore_index=True)

    
    # =================================================
    # 整併 DATASET
    # =================================================
    if visit_df is not None and not result.empty:

        visit_map_df = (
            visit_df[["STAT_DATASET", "FORM_NAME"]]
            .dropna()
            .drop_duplicates()
            .copy()
        )
        
        visit_map_df["FORM_NAME"] = visit_map_df["FORM_NAME"].astype(str).str.strip().str.upper().apply(normalize_form_name)
        visit_map_df["STAT_DATASET"] = visit_map_df["STAT_DATASET"].astype(str).str.strip().str.upper()


        result["FORM_NORM"] = result["FORM"].apply(normalize_form_name)

        result = result.merge(
            visit_map_df,
            left_on="FORM_NORM",
            right_on="FORM_NAME",
            how="left"
        )

        result.drop(columns=["FORM_NAME", "FORM_NORM"], inplace=True)

        cols = result.columns.tolist()

        if "STAT_DATASET" in cols and "FORM" in cols:
            cols.remove("STAT_DATASET")
            form_idx = cols.index("FORM")
            cols.insert(form_idx, "STAT_DATASET")

            result = result[cols]

    return result





# =========================================================
# Condition tokenization / clean
# =========================================================

# reference range 關鍵字
REFERENCE_RANGE_TOKENS = {
    "ULN", "LLN", "URN", "LRN",
    "NORMAL RANGE",
    "REFERENCE RANGE",

    "UPPER LIMIT", "LOWER LIMIT",
    "UPPER LIMIT OF NORMAL", "LOWER LIMIT OF NORMAL",
    "LOWER LIMIT NORMAL", "UPPER LIMIT NORMAL",
}

# unit token 的唯一維護來源
UNIT_TOKEN_FRAGMENTS = {
    "L", "DL", "ML",
    "G", "MG", "MMOL", "UMOL",
    "U", "IU", 
    "KG", "CM", "MMHG",
}

# Reserved tokens = parser 層，不進 UNKNOWN_VAR
RESERVED_TOKENS = {
    "AND", "OR", "NOT", "IN", "IS", "MISSING", "TRUE", "FALSE",
    "EQ", "NE", "LT", "LE", "GT", "GE", "GTE", "LTE",
    "MAX", "MIN",

    # function
    "YEAR",

    # common non-variable words
    "PD", "NA", "NAN", "NONE",
    "DATE", "TIME",

    # reference range tokens 不進 UNKNOWN_VAR
    *{x for x in REFERENCE_RANGE_TOKENS if " " not in x},  # 單字的才進token

    # unit fragments 從唯一來源展開
    *UNIT_TOKEN_FRAGMENTS,

}

POSITION_SUFFIXES = {"PREV", "PREVIOUS", "NEXT", "FIRST", "LAST"}

# semantic pattern：用 UNIT_TOKEN_FRAGMENTS 產生
_UNIT_ALT = "|".join(sorted(UNIT_TOKEN_FRAGMENTS, key=len, reverse=True))


UNIT_DEPENDENT_PATTERNS = [
    # scientific notation / scale
    r"10\^",
    r"10\*\*",
    r"×10",
    r"\bX10\b",
    r"\b10\s*\^\s*\d+\b",

    # 例如 10^9/L, mg/dL, mmol/L
    rf"/\s*({_UNIT_ALT})\b",

    # 例如 G/L, MG/DL, MMOL/L
    rf"\b({_UNIT_ALT})\s*/\s*({_UNIT_ALT})\b",
]

ITEM_KEY_SUFFIX_PATTERN = r"(TEST|TESTCD|ITEM|ITEMCD|PARAM|PARAMCD)$"

ITEM_RESULT_SUFFIX_PATTERNS = [
    r"ORRES$",
    r"STRESN$",
    r"STRESC$",
    r"STAT$",
    r"CLISIG$",
    r"RES$",
    r"RESULT$",
    r"ND$",
    r"NOR$",
]



def clean_condition_text(condition):
    if condition is None or pd.isna(condition):
        return ""
    s = str(condition)
    s = s.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
    s = s.replace("＝", "=").replace("＜", "<").replace("＞", ">")
    s = s.replace("＋", "+").replace("－", "-").replace("＊", "*").replace("×", "*")
    s = s.replace("：", ":").replace("，", ",").replace("～", "~")

    # Remove non-executable instructions after To SAS/To DM/Note labels
    s = re.split(r"(?is)\bto\s*(dm|sas|sasp)(?:[\s_].*)?$", s)[0]
    s = re.split(r"(?is)\b(note|message|註)\s*[:：].*$", s)[0]
    s = re.sub(r"(?m)\*(?!\s*\d).*$", "", s)
    s = re.sub(r"[ \t\f\v]+", " ", s)
    s = re.sub(r"\n\s*\n+", "\n", s)

    return s.strip()


# 抓動態參數 (例如: X=2, 3, 4...)，先標記，未展開
def parse_param_map_from_condition(condition):
    if condition is None or pd.isna(condition):
        return {}
    param_map = {}
    for line in re.split(r"[\r\n]+", str(condition)):
        line = line.strip().replace("...", "").replace("…", "")
        m = re.fullmatch(r"([A-Z])\s*=\s*([A-Za-z0-9~,_\-\s]+)", line, flags=re.I)
        if not m:
            continue
        key = m.group(1).upper()
        vals = [x.strip().upper() for x in m.group(2).split(",") if x.strip()]
        if vals:
            param_map[key] = vals
    return param_map


# 拆condition 裡的 labels 例如 C1D1:, C2D1, CXD1:
def split_condition_into_visit_blocks(condition, raw_spec=None):

    if condition is None or pd.isna(condition):
        return []
    
    # =====================================================
    # 1. 從raw_spec的visit list找合法的visit
    # =====================================================
    visit_df = None

    if isinstance(raw_spec, dict):
        if raw_spec.get("visit") is not None:
            visit_df = raw_spec.get("visit")
        elif raw_spec.get("domain_visit") is not None:
            visit_df = raw_spec.get("domain_visit")

    visit_set = set()

    if visit_df is not None and not visit_df.empty and "VISIT_DVP" in visit_df.columns:
        visit_set = set(
            visit_df["VISIT_DVP"]
            .dropna()
            .astype(str)
            .str.strip()
            .str.upper()
            .tolist()
        )

    # 給 expand_visit_scope_text 使用，確保它吃得到 visit
    raw_spec_for_visit = raw_spec

    if isinstance(raw_spec, dict) and visit_df is not None:
        raw_spec_for_visit = dict(raw_spec)
        raw_spec_for_visit["visit"] = visit_df

    # =====================================================
    # 2. Clean first
    # =====================================================
    s = clean_condition_text(condition)

    s = (
        str(s)
        .replace("：", ":")
        .replace("，", ",")
        .replace("～", "~")
    )

    # 保險：clean_condition_text 應已處理，但這裡再保留一次
    s = re.split(
        r"(?i)\bto\s*[_\s,]*(dm|sas|sasp)\b",
        s
    )[0]

    lines = [x.rstrip() for x in re.split(r"[\r\n]+", s)]

    blocks = []
    cur_scope = []
    cur_invalid_labels = []
    cur_lines = []

    # =====================================================
    # 3. Flush current block
    # =====================================================
    def flush():
        nonlocal cur_lines, cur_scope, cur_invalid_labels

        expr = "\n".join(cur_lines).strip()

        if not expr:
            cur_lines = []
            cur_invalid_labels = []
            return

        # valid visit labels：一個 visit 一個 block
        if cur_scope:
            for vs in cur_scope:
                blocks.append({
                    "BLOCK_VISIT_SCOPE": [vs],
                    "INVALID_VISIT_LABELS": cur_invalid_labels.copy(),
                    "CONDITION_BLOCK": expr
                })

        # invalid label 情況：仍保留 block，讓 Step1 UI 可以 QC
        elif cur_invalid_labels:
            blocks.append({
                "BLOCK_VISIT_SCOPE": [],
                "INVALID_VISIT_LABELS": cur_invalid_labels.copy(),
                "CONDITION_BLOCK": expr
            })

        # no label condition
        else:
            blocks.append({
                "BLOCK_VISIT_SCOPE": [],
                "INVALID_VISIT_LABELS": [],
                "CONDITION_BLOCK": expr
            })

        cur_lines = []
        cur_invalid_labels = []

    # =====================================================
    # 4. Parse lines
    # =====================================================
    for line in lines:

        ls = str(line).strip()

        if not ls:
            continue

        ls_norm = (
            ls.upper()
            .replace("：", ":")
            .replace("，", ",")
            .replace("～", "~")
        )


        # =====================================================
        # Skip section marker (例如: (1)、(2)...)
        # =====================================================
        if re.fullmatch(r"\(?\d+\)?\.?", ls_norm):
            continue

        # =====================================================
        # Skip parameter line:
        #   X = 2, 3, 4...
        # =====================================================
        if re.fullmatch(
            r"[A-Z]\s*=\s*[A-Za-z0-9~,_\-\s\.…]+",
            ls_norm,
            flags=re.IGNORECASE
        ):
            continue

        # =====================================================
        # Visit label detection
        # 僅接受 raw_spec VISIT_DVP 裡存在的 visit
        # =====================================================
        if ls_norm.endswith(":"):

            label_part = ls_norm.rstrip(":").strip()

            # 因為可能是：
            #   C2D1, CXD1
            #   C1D1~CXD15
            # 所以直接使用 expand_visit_scope_text()
            expanded_scope = expand_visit_scope_text(
                label_part,
                raw_spec=raw_spec_for_visit
            )

            expanded_scope = [
                str(v).strip().upper()
                for v in expanded_scope
                if str(v).strip()
            ]

            # 只接受 schema 內存在的 VISIT_DVP
            if visit_set:
                valid_scope = [
                    v for v in expanded_scope
                    if v in visit_set
                ]

                invalid_scope = [
                    v for v in expanded_scope
                    if v not in visit_set
                ]

            else:
                # 如果沒有 raw_spec，就不主動接受 label
                # 避免 regex 誤判造成錯誤 block
                valid_scope = []
                invalid_scope = expanded_scope

            # 如果這行確實像 visit label，但有 valid 或 invalid，
            # 都應該切 block，不要把 label 丟進 condition
            if valid_scope or invalid_scope:

                flush()

                cur_scope = valid_scope
                cur_invalid_labels = invalid_scope

                continue
        

        ls_low = ls.lower()

        # case 1：單獨一行 "or" / "and"
        if ls_low in ("or", "and"):
            if cur_lines:
                cur_lines[-1] = cur_lines[-1] + " " + ls
            continue

        # case 2：上一行接 or（例如 "... 0.95)) or"）
        if cur_lines and cur_lines[-1].strip().lower().endswith(("or", "and")):
            cur_lines[-1] = cur_lines[-1] + " " + ls
            continue

        # case 3：正常
        cur_lines.append(ls)



    # =====================================================
    # 5. Final flush
    # =====================================================
    flush()

    # =====================================================
    # 6. Fallback: no explicit label
    # =====================================================
    if not blocks:

        expr = clean_condition_text(condition)

        if expr:
            blocks = [{
                "BLOCK_VISIT_SCOPE": [],
                "INVALID_VISIT_LABELS": [],
                "CONDITION_BLOCK": expr
            }]

    return blocks



# 抓condition 中的所有變數 token
def find_identifiers_in_condition(condition):
    if condition is None or pd.isna(condition):
        return []
    s = clean_condition_text(condition)
    s = re.sub(r'"[^"]*"|\'[^\']*\'', " ", s)
    tokens = re.findall(r"\b[A-Za-z_][A-Za-z0-9_]*\b", s)
    out = []
    for t in tokens:
        tu = t.upper()
        if tu in RESERVED_TOKENS:
            continue
        if re.fullmatch(r"C(?:\d+|X)D\d+", tu):
            continue
        if tu in POSITION_SUFFIXES:
            continue
        out.append(tu)
    return list(dict.fromkeys(out))


# 解析 token 類型，決定後面 merge 行為
def normalize_token_meta(token):
    """Return metadata for token: base_var, suffix type, raw token."""
    raw = str(token).strip().upper()
    if "_" not in raw:
        return {"RAW_TOKEN": raw, "BASE_VAR": raw, "TOKEN_TYPE": "EXACT", "SUFFIX": None}
    base, suffix = raw.rsplit("_", 1)
    if suffix in POSITION_SUFFIXES:
        return {"RAW_TOKEN": raw, "BASE_VAR": base, "TOKEN_TYPE": "POSITION_SUFFIX", "SUFFIX": suffix}
    if suffix == "X" or re.fullmatch(r"CXD\d+", suffix):
        return {"RAW_TOKEN": raw, "BASE_VAR": base, "TOKEN_TYPE": "DYNAMIC_VISIT_SUFFIX", "SUFFIX": suffix}
    if re.fullmatch(r"[A-Z0-9]+", suffix):
        return {"RAW_TOKEN": raw, "BASE_VAR": base, "TOKEN_TYPE": "FIXED_VISIT_SUFFIX", "SUFFIX": suffix}
    return {"RAW_TOKEN": raw, "BASE_VAR": raw, "TOKEN_TYPE": "EXACT", "SUFFIX": None}





# =========================================================
# Visit scope helpers
# =========================================================

def expand_visit_scope_text(visit_text, raw_spec=None):
    if visit_text is None or pd.isna(visit_text):
        return []
    s = str(visit_text).strip()
    if s == "" or s.upper() == "ALL":
        return []
    s = s.replace("，", ",").replace("～", "~")

    visit_df = raw_spec.get("visit") if isinstance(raw_spec, dict) else None
    visit_master = pd.DataFrame()
    if visit_df is not None and not visit_df.empty and "VISIT_DVP" in visit_df.columns:
        visit_master = normalize_upper_cols(visit_df)
        visit_master["VISIT_DVP"] = visit_master["VISIT_DVP"].astype(str).str.strip().str.upper()
        if "VISIT_ORDER" not in visit_master.columns:
            visit_master["VISIT_ORDER"] = range(1, len(visit_master) + 1)

    def norm_one(tok):
        t = str(tok).strip().upper()
        if not t:
            return []
        fallback = {
            "SCREENING": "S", "SCREEN": "S", "ENROLLMENT": "EN",
            "END OF TREATMENT": "EOT", "END OF STUDY": "EOS", "UNSCHEDULED": "UN", "UNS": "UN",
        }
        t = fallback.get(t, t)
        if visit_master.empty:
            return [t]
        hit = visit_master[visit_master["VISIT_DVP"] == t]
        return hit["VISIT_DVP"].tolist() if not hit.empty else [t]

    def expand_range(a, b):
        aa, bb = norm_one(a), norm_one(b)
        if visit_master.empty or not aa or not bb:
            return aa + bb
        a0, b0 = aa[0], bb[0]
        if a0 not in set(visit_master["VISIT_DVP"]) or b0 not in set(visit_master["VISIT_DVP"]):
            return [a0, b0]
        ao = visit_master.loc[visit_master["VISIT_DVP"] == a0, "VISIT_ORDER"].iloc[0]
        bo = visit_master.loc[visit_master["VISIT_DVP"] == b0, "VISIT_ORDER"].iloc[0]
        lo, hi = min(ao, bo), max(ao, bo)
        return visit_master[(visit_master["VISIT_ORDER"] >= lo) & (visit_master["VISIT_ORDER"] <= hi)]["VISIT_DVP"].tolist()

    out = []
    for part in [x.strip() for x in s.split(",") if x.strip()]:
        if "~" in part:
            p = [x.strip() for x in part.split("~") if x.strip()]
            out.extend(expand_range(p[0], p[1]) if len(p) == 2 else norm_one(part))
        else:
            out.extend(norm_one(part))
    return list(dict.fromkeys(out))


def add_dynamic_visit_context_columns(df):
    if df is None:
        return pd.DataFrame()
    if df.empty:
        return df.copy()
    out = normalize_upper_cols(df)
    for c in ["VISIT_DVP", "VISIT"]:
        if c in out.columns:
            out[c] = out[c].astype(str).str.strip().str.upper()
    if "VISITNUM" in out.columns:
        out["VISITNUM"] = pd.to_numeric(out["VISITNUM"], errors="coerce")

    def cycle_from_dvp(v):
        m = re.fullmatch(r"C(\d+)D\d+", str(v).strip().upper()) if pd.notna(v) else None
        return int(m.group(1)) if m else pd.NA
    def day_from_dvp(v):
        m = re.fullmatch(r"C(?:\d+|X)D(\d+)", str(v).strip().upper()) if pd.notna(v) else None
        return int(m.group(1)) if m else pd.NA
    def cycle_from_visit(v):
        if pd.isna(v): return pd.NA
        s = str(v).upper()
        m = re.search(r"CYCLE\s+(\d+)", s) or re.search(r"\bC(\d+)D\d+\b", s)
        return int(m.group(1)) if m else pd.NA
    def day_from_visit(v):
        if pd.isna(v): return pd.NA
        s = str(v).upper()
        m = re.search(r"DAY\s+(\d+)", s) or re.search(r"\bC(?:\d+|X)D(\d+)\b", s)
        return int(m.group(1)) if m else pd.NA
    def cycle_from_num(v):
        try: n = int(float(v))
        except Exception: return pd.NA
        return int(n // 10000) if n >= 10000 and (n // 10000) > 0 else pd.NA
    def day_from_num(v):
        try: n = int(float(v))
        except Exception: return pd.NA
        d = (n % 10000) // 100 if n >= 10000 else 0
        return int(d) if d > 0 else pd.NA

    out["X"] = pd.NA
    out["VISIT_DAY"] = pd.NA
    if "VISIT_DVP" in out.columns:
        out["X"] = out["VISIT_DVP"].apply(cycle_from_dvp)
        out["VISIT_DAY"] = out["VISIT_DVP"].apply(day_from_dvp)
    if "VISIT" in out.columns:
        out["X"] = out["X"].where(out["X"].notna(), out["VISIT"].apply(cycle_from_visit))
        out["VISIT_DAY"] = out["VISIT_DAY"].where(out["VISIT_DAY"].notna(), out["VISIT"].apply(day_from_visit))
    if "VISITNUM" in out.columns:
        out["X"] = out["X"].where(out["X"].notna(), out["VISITNUM"].apply(cycle_from_num))
        out["VISIT_DAY"] = out["VISIT_DAY"].where(out["VISIT_DAY"].notna(), out["VISITNUM"].apply(day_from_num))
    out["X"] = pd.to_numeric(out["X"], errors="coerce")
    out["VISIT_DAY"] = pd.to_numeric(out["VISIT_DAY"], errors="coerce")
    return out


def attach_visit_order(df, raw_spec):
    if df is None:
        return pd.DataFrame()
    out = normalize_upper_cols(df)
    if "VISIT_DVP" not in out.columns:
        return out
    visit_df = raw_spec.get("visit") if isinstance(raw_spec, dict) else None
    if visit_df is None or visit_df.empty or "VISIT_DVP" not in visit_df.columns:
        return out
    v = normalize_upper_cols(visit_df)
    if "VISIT_ORDER" not in v.columns:
        if "VISITNUM" in v.columns:
            v["VISIT_ORDER"] = v["VISITNUM"]
        else:
            v["VISIT_ORDER"] = range(1, len(v) + 1)
    lk = v[["VISIT_DVP", "VISIT_ORDER"]].drop_duplicates()
    out["VISIT_DVP"] = out["VISIT_DVP"].astype(str).str.strip().str.upper()
    return out.merge(lk, on="VISIT_DVP", how="left") if "VISIT_ORDER" not in out.columns else out


# 根據單一 rule/block 的 visit scope filter eval_df。
def filter_eval_df_by_visit_scope(eval_df, review_row):
    if eval_df is None or eval_df.empty:
        return eval_df
    
    out = add_dynamic_visit_context_columns(eval_df)
    
    if review_row is None:
        return out

    # 1. 最優先：EFFECTIVE_VISIT
    effective_visit = review_row.get("EFFECTIVE_VISIT", None)

    scopes = []

    if effective_visit is not None and pd.notna(effective_visit) and str(effective_visit).strip() != "":
        scopes = [str(effective_visit).strip().upper()]

    else:
        # 2. fallback scopes
        for key in [
            "EFFECTIVE_VISIT_SCOPE",
            "BLOCK_VISIT_SCOPE",
            "DVP_VISIT_SCOPE"
        ]:
            val = review_row.get(key, [])

            if isinstance(val, list) and len(val) > 0:
                scopes = [
                    str(x).strip().upper()
                    for x in val
                    if str(x).strip()
                ]
                break

    if not scopes:
        return out

    visit_col = None

    if "VISIT_DVP" in out.columns:
        visit_col = "VISIT_DVP"
    elif "VISIT" in out.columns:
        visit_col = "VISIT"

    if visit_col is None:
        return out

    out[visit_col] = (out[visit_col].astype(str).str.strip().str.upper())

    masks = []

    for scope in scopes:

        scope = str(scope).strip().upper()
        # =====================================================
        # CXD1 / CXD15 (這裡只抓 literal CXD1，不再把 C2D1 / C3D1 混進來)
        # =====================================================
        masks.append(out[visit_col].eq(scope))

    if not masks:
        return out

    final_mask = masks[0]

    for m in masks[1:]:
        final_mask = final_mask | m

    return out[final_mask].copy()




# ===================================================================================================================================================
# STEP 1 review table
# ===================================================================================================================================================
# 把 tokens 輸出分類為 target / ref / row_op / unknown
def build_ref_specs_from_tokens(tokens, stat_dataset, raw_spec, context_text=""):
    """
    重要：
      若同一 VARIABLE 出現在多個 STAT_DATASET，
      先用 context_text 比對 schema label 來 resolve。

    例：
      DSSTDAT 同時在 DS / EN
      context_text 有 "Informed Consent Date"
      schema label:
        DS.DSSTDAT = Date of Study Completion
        EN.DSSTDAT = Informed Consent Date
      => resolve to EN
    """

    var_to_stat = build_var_to_stat_map(raw_spec)
    stat_to_vars = build_stat_to_var_map(raw_spec)

    target_vars, refs, row_ops, unknown = [], [], [], []

    valid_vars = get_valid_var_set(raw_spec)

    stat_dataset = str(stat_dataset).strip().upper()
    target_var_set = stat_to_vars.get(stat_dataset, set())

    context_upper = str(context_text or "").upper()

    # =====================================================
    # Prepare variable label lookup from raw_spec["variable"]
    # =====================================================
    variable_df = raw_spec.get("variable") if isinstance(raw_spec, dict) else pd.DataFrame()

    var_meta_df = pd.DataFrame()

    if variable_df is not None and not variable_df.empty:
        var_meta_df = normalize_upper_cols(variable_df)

    label_cols = []

    if not var_meta_df.empty:
        for c in [
            "LABEL",
            "VARIABLE_LABEL",
            "VAR_LABEL",
            "FIELD_LABEL",
            "PROMPT",
            "QUESTION",
        ]:
            if c in var_meta_df.columns:
                label_cols.append(c)

    seen_refs = set()

    for tok in tokens:

        meta = normalize_token_meta(tok)

        raw_token = meta["RAW_TOKEN"]
        base_var = meta["BASE_VAR"]
        ttype = meta["TOKEN_TYPE"]
        suffix = meta["SUFFIX"]

        if base_var in RESERVED_TOKENS or raw_token in RESERVED_TOKENS:
            continue

        if base_var not in valid_vars and base_var != "X":
            unknown.append(raw_token)
            continue

        if raw_token == "X" or base_var == "X":
            continue

        # =====================================================
        # Row ops: _PREV / _NEXT / _FIRST / _LAST
        # =====================================================
        if ttype == "POSITION_SUFFIX":

            source_stat_dataset = infer_row_op_source_stat_dataset(
                source_var=base_var,
                target_stat_dataset=stat_dataset,
                raw_spec=raw_spec
            )

            row_ops.append({
                "SOURCE_VAR": base_var,
                "SOURCE_STAT_DATASET": source_stat_dataset,
                "TARGET_STAT_DATASET": stat_dataset,
                "DERIVED_VAR": raw_token,
                "OP": suffix
            })

            continue

        candidate_stats = sorted(var_to_stat.get(base_var, []))
        base_in_target = base_var in target_var_set

        # =====================================================
        # Same dataset target var
        # =====================================================
        if ttype == "EXACT" and base_in_target:
            target_vars.append(base_var)
            continue

        # =====================================================
        # Resolve ref dataset
        # =====================================================
        ref_stat = ""
        ambiguous_candidates = []

        if base_in_target and ttype in ["FIXED_VISIT_SUFFIX", "DYNAMIC_VISIT_SUFFIX"]:
            ref_stat = stat_dataset

        elif len(candidate_stats) == 1:
            ref_stat = candidate_stats[0]

        elif len(candidate_stats) > 1:

            # ---------------------------------------------
            # Try label match using context text
            # ---------------------------------------------
            matched_stats = []

            if not var_meta_df.empty and label_cols:
                tmp = var_meta_df[
                    var_meta_df["VARIABLE"].astype(str).str.strip().str.upper().eq(base_var)
                ].copy()

                if "STAT_DATASET" in tmp.columns:

                    for _, mr in tmp.iterrows():

                        cand_stat = str(mr.get("STAT_DATASET", "")).strip().upper()

                        if cand_stat not in candidate_stats:
                            continue

                        for lc in label_cols:

                            label = str(mr.get(lc, "")).strip().upper()

                            if not label or label in ["NAN", "NONE"]:
                                continue

                            if label in context_upper:
                                matched_stats.append(cand_stat)
                                break

            matched_stats = list(dict.fromkeys(matched_stats))

            if len(matched_stats) == 1:
                ref_stat = matched_stats[0]

            else:
                # 無法唯一判斷，不要默默選第一個
                ambiguous_candidates = candidate_stats

        # =====================================================
        # Build ref spec
        # =====================================================
        if ttype in ["FIXED_VISIT_SUFFIX", "DYNAMIC_VISIT_SUFFIX"]:

            join_type = (
                "DYNAMIC_CYCLE"
                if (suffix == "X" or re.fullmatch(r"CXD\d+", str(suffix)))
                else "FIXED_VISIT"
            )

            ref = {
                "RAW_TOKEN": raw_token,
                "BASE_VAR": base_var,
                "REF_STAT_DATASET": ref_stat,
                "JOIN_TYPE": join_type,
                "VISIT_SUFFIX": suffix,
                "PARAM": "X" if join_type == "DYNAMIC_CYCLE" else None,
            }

        else:

            ref = {
                "RAW_TOKEN": raw_token,
                "BASE_VAR": base_var,
                "REF_STAT_DATASET": ref_stat,
                "JOIN_TYPE": "AUTO",
                "VISIT_SUFFIX": None,
                "PARAM": None,
            }

        if ambiguous_candidates:
            ref["AMBIGUOUS_REF_DATASETS"] = ambiguous_candidates
            ref["JOIN_TYPE"] = "AMBIGUOUS"
            ref["REF_STAT_DATASET"] = ""

        key = (
            ref.get("RAW_TOKEN"),
            ref.get("BASE_VAR"),
            ref.get("REF_STAT_DATASET"),
            ref.get("JOIN_TYPE"),
            str(ref.get("VISIT_SUFFIX"))
        )

        if key not in seen_refs:
            refs.append(ref)
            seen_refs.add(key)

    return (
        list(dict.fromkeys(target_vars)),
        refs,
        row_ops,
        list(dict.fromkeys(unknown))
    )



# 判斷 Rule 的類型
def infer_rule_type_from_specs(refs, row_ops, params, unknown):
    if unknown:
        return "UNRESOLVED"
    if row_ops:
        return "WITH_ROW_OP"
    if any(r.get("JOIN_TYPE") == "DYNAMIC_CYCLE" for r in refs) or params:
        return "DYNAMIC_PARAM"
    if refs:
        return "CROSS_DATASET"
    return "SINGLE_DATASET"



def build_dvp_review_df(dvp_df, raw_spec):
    """
    Step 1 主表: DVP Review / Execution Spec

    每一列 = 一個 rule block。

    主要輸出：
      - RULE / STAT_DATASET / BLOCK_ID
      - DVP_VISIT_SCOPE / BLOCK_VISIT_SCOPE / EFFECTIVE_VISIT_SCOPE
      - INVALID_VISIT_LABELS
      - TARGET_VARS / REF_SPECS / REF_DATASETS / REF_VARS
      - ROW_OPS
      - UNKNOWN_VARS
      - RULE_TYPE
      - EXECUTION_CLASS
      - NEED_REFERENCE_RANGE
      - NEED_UNIT_STANDARDIZATION
      - BLOCKING_REASON
      - READY_FOR_STEP2
    """

    if dvp_df is None or dvp_df.empty:
        return pd.DataFrame()

    rows = []

    for _, r in dvp_df.iterrows():

        # =====================================================
        # Basic fields
        # =====================================================
        condition = r.get("CONDITION", "")

        params = parse_param_map_from_condition(condition)

        # DVP Visit/Folder scope
        dvp_scope = expand_visit_scope_text(
            r.get("VISIT", ""),
            raw_spec=raw_spec
        )

        dvp_scope = [
            str(x).strip().upper()
            for x in dvp_scope
            if str(x).strip()
        ]

        # =====================================================
        # Split condition into visit blocks
        # 注意：
        #   這裡要傳 raw_spec，因為 block visit label 必須用 schema VISIT_DVP 驗證
        # =====================================================
        blocks = split_condition_into_visit_blocks(
            condition,
            raw_spec=raw_spec
        )

        if not blocks:
            blocks = [{
                "BLOCK_VISIT_SCOPE": [],
                "INVALID_VISIT_LABELS": [],
                "CONDITION_BLOCK": clean_condition_text(condition)
            }]

        # =====================================================
        # STAT_DATASET
        # parse_dvp_file 已經會從 FORM mapping STAT_DATASET
        # 所以這裡優先使用 STAT_DATASET
        # =====================================================
        stat_dataset = str(
            r.get(
                "STAT_DATASET",
                r.get("DATASET", r.get("FORM", ""))
            )
        ).strip().upper()

        form = r.get("FORM", "")
        field = r.get("FIELD", "")
        message = r.get("MESSAGE", r.get("NOTE", ""))

        # =====================================================
        # Per block
        # =====================================================
        for i, b in enumerate(blocks, start=1):

            block_expr = clean_condition_text(
                b.get("CONDITION_BLOCK", "")
            )

            # 如果 clean 後是空，代表這個 block 其實只是 Note / To DM / comment
            if not str(block_expr).strip():
                continue

            block_scope = b.get("BLOCK_VISIT_SCOPE", []) or []
            block_scope = [
                str(x).strip().upper()
                for x in block_scope
                if str(x).strip()
            ]

            invalid_visit_labels = b.get("INVALID_VISIT_LABELS", []) or []
            invalid_visit_labels = [
                str(x).strip().upper()
                for x in invalid_visit_labels
                if str(x).strip()
            ]

            # =====================================================
            # Tokens
            # =====================================================
            tokens = find_identifiers_in_condition(block_expr)

            # =====================================================
            # target / ref / row_ops / unknown
            # =====================================================
            context_text = " ".join([
                str(condition or ""),
                str(message or ""),
                str(field or ""),
                str(form or ""),
            ])

            target_vars, refs, row_ops, unknown = build_ref_specs_from_tokens(
                tokens=tokens,
                stat_dataset=stat_dataset,
                raw_spec=raw_spec,
                context_text=context_text
            )

            ref_datasets = sorted({
                x.get("REF_STAT_DATASET")
                for x in refs
                if x.get("REF_STAT_DATASET")
            })

            ref_vars = list(dict.fromkeys([
                x.get("RAW_TOKEN")
                for x in refs
                if x.get("RAW_TOKEN")
            ]))

            ambiguous_refs = [
                x for x in refs
                if x.get("AMBIGUOUS_REF_DATASETS")
            ]

            # =====================================================
            # Rule type
            # =====================================================
            rule_type = infer_rule_type_from_specs(
                refs=refs,
                row_ops=row_ops,
                params=params,
                unknown=unknown
            )

            # =====================================================
            # Execution feasibility flags
            # =====================================================
            block_upper = str(block_expr).upper()

            # Reference range，例如 ULN / LLN / reference range
            needs_reference_range = any(
                token in block_upper
                for token in REFERENCE_RANGE_TOKENS
            )

            # Unit / scale，例如 10^9/L, g/L, mg/dL
            needs_unit_standardization = any(
                re.search(pattern, block_upper, flags=re.IGNORECASE)
                for pattern in UNIT_DEPENDENT_PATTERNS
            )

            blocking_reasons = []

            if unknown:
                blocking_reasons.append(
                    "Condition contains variables not found in schema."
                )

            if invalid_visit_labels:
                blocking_reasons.append(
                    "Condition block contains visit labels not found in schema VISIT_DVP."
                )

            if not stat_dataset:
                blocking_reasons.append(
                    "STAT_DATASET is missing."
                )

            if needs_reference_range:
                blocking_reasons.append(
                    "Condition depends on LLN/ULN/reference range not available in raw data."
                )

            if needs_unit_standardization:
                blocking_reasons.append(
                    "Condition contains explicit unit/scale; raw data may need unit conversion before execution."
                )

            if ambiguous_refs:
                blocking_reasons.append(
                    "Condition contains variables that exist in multiple STAT_DATASETs and cannot be resolved by label/context."
                )


            # =====================================================
            # Execution class priority
            # =====================================================
            if unknown:
                execution_class = "UNRESOLVED_SCHEMA"

            elif ambiguous_refs:
                execution_class = "AMBIGUOUS_REF_DATASET"

            elif invalid_visit_labels:
                execution_class = "INVALID_VISIT_LABEL"

            elif needs_reference_range:
                execution_class = "NEED_REFERENCE_RANGE"

            elif needs_unit_standardization:
                execution_class = "NEED_UNIT_STANDARDIZATION"

            else:
                execution_class = "EXECUTABLE"

            ready = (
                len(unknown) == 0
                and len(ambiguous_refs) == 0
                and len(invalid_visit_labels) == 0
                and stat_dataset != ""
                and execution_class == "EXECUTABLE"
            )

            # =====================================================
            # Effective visit scope
            # =====================================================
            effective_visit_scope = block_scope if block_scope else dvp_scope

            effective_visit_scope = [
                str(x).strip().upper()
                for x in effective_visit_scope
                if str(x).strip()
            ]

            # =====================================================
            # Visit expansion policy
            #   只有 executable 的 rule/block 才展開 visit。
            #   如果是 UNRESOLVED_SCHEMA / NEED_UNIT / NEED_RANGE / INVALID_VISIT_LABEL，不展開
            #   但仍保留 EFFECTIVE_VISIT_SCOPE list，方便 UI review。
            # =====================================================
            if ready:
                expanded_effective_visits = (
                    effective_visit_scope
                    if effective_visit_scope
                    else [None]
                )
                visit_expanded = True
            else:
                expanded_effective_visits = [None]
                visit_expanded = False

            for effective_visit in expanded_effective_visits:

                rows.append({
                    # -------------------------
                    # identity
                    # -------------------------
                    "RULE": r.get("RULE", ""),
                    "STAT_DATASET": stat_dataset,
                    "FORM": form,
                    "FIELD": field,
                    "MESSAGE": message,
                    "PD": r.get("PD", ""),

                    # -------------------------
                    # visit
                    # -------------------------
                    "VISIT_RAW": r.get("VISIT", ""),
                    "DVP_VISIT_SCOPE": dvp_scope,
                    "BLOCK_ID": i,
                    "BLOCK_VISIT_SCOPE": block_scope,
                    "INVALID_VISIT_LABELS": invalid_visit_labels,

                    "EFFECTIVE_VISIT": effective_visit,

                    "EFFECTIVE_VISIT_SCOPE": (
                        [effective_visit]
                        if effective_visit is not None
                        else effective_visit_scope
                    ),

                    # UI/debug 用
                    "VISIT_EXPANDED": bool(visit_expanded),

                    # -------------------------
                    # condition
                    # -------------------------
                    "CONDITION": condition,
                    "CONDITION_BLOCK": block_expr,

                    # -------------------------
                    # parsed vars
                    # -------------------------
                    "RULE_VAR_TOKENS": tokens,
                    "TARGET_VARS": target_vars,
                    "REF_SPECS": refs,
                    "REF_DATASETS": ref_datasets,
                    "REF_VARS": ref_vars,
                    "ROW_OPS": row_ops,
                    "UNKNOWN_VARS": unknown,
                    "AMBIGUOUS_REFS": ambiguous_refs,

                    # -------------------------
                    # params
                    # -------------------------
                    "PARAMS": params,
                    "PARAM_STRATEGY": (
                        "DYNAMIC_FROM_DATA"
                        if params or any(re.search(r"_CXD\d+", x) for x in tokens)
                        else ""
                    ),

                    # -------------------------
                    # classification
                    # -------------------------
                    "RULE_TYPE": rule_type,
                    "EXECUTION_CLASS": execution_class,
                    "NEED_REFERENCE_RANGE": bool(needs_reference_range),
                    "NEED_UNIT_STANDARDIZATION": bool(needs_unit_standardization),
                    "BLOCKING_REASON": " | ".join(blocking_reasons),

                    # -------------------------
                    # final readiness
                    # -------------------------
                    "READY_FOR_STEP2": bool(ready),
                })


    return pd.DataFrame(rows)





# =========================================================
# Ref merge engine v2
# =========================================================

# 判斷 ref 是：SUBJECT 還是 VISIT level
def resolve_auto_ref_grain(ref_df, base_var):
    if ref_df is None or ref_df.empty:
        return "SUBJECT", {"REASON": "EMPTY_REF"}
    df = normalize_upper_cols(ref_df)
    base_var = str(base_var).strip().upper()
    subj, visit = get_subject_key(df), get_visit_key(df)
    if subj is None:
        return "SUBJECT", {"REASON": "NO_SUBJECT_KEY"}
    if visit is None:
        return "SUBJECT", {"REASON": "NO_VISIT_KEY"}
    if base_var not in df.columns:
        return "SUBJECT", {"REASON": "BASE_VAR_NOT_FOUND"}
    tmp = df[[subj, visit, base_var]].copy()
    tmp[subj] = normalize_join_value_series(tmp[subj])
    tmp[visit] = normalize_join_value_series(tmp[visit])
    tmp[base_var] = tmp[base_var].replace("", pd.NA)
    tmp = tmp[tmp[subj].notna() & tmp[base_var].notna()].copy()
    if tmp.empty:
        return "SUBJECT", {"REASON": "NO_NON_MISSING_VALUES"}
    visit_counts = tmp.groupby(subj)[visit].nunique(dropna=True)
    value_counts = tmp.groupby(subj)[base_var].nunique(dropna=True)
    max_visit_n = int(visit_counts.max()) if not visit_counts.empty and pd.notna(visit_counts.max()) else 0
    max_value_n = int(value_counts.max()) if not value_counts.empty and pd.notna(value_counts.max()) else 0
    is_visit = max_visit_n > 1 and max_value_n > 1
    return ("VISIT" if is_visit else "SUBJECT"), {"MAX_VISIT_N_PER_SUBJECT": max_visit_n, "MAX_VALUE_N_PER_SUBJECT": max_value_n, "IS_VISIT_LEVEL": bool(is_visit)}


# 用 SUBJID 合併
def materialize_subject_ref(target_df, ref_df, raw_token, base_var, ref_obj):
    out, ref = normalize_upper_cols(target_df), normalize_upper_cols(ref_df)
    raw_token, base_var = str(raw_token).upper(), str(base_var).upper()
    subj_t, subj_r = get_subject_key(out), get_subject_key(ref)
    if subj_t is None or subj_r is None or base_var not in ref.columns:
        out[raw_token] = pd.NA
        ref_obj.setdefault("MERGE_DEBUG", {})["ERROR"] = "SUBJECT_REF_KEY_OR_VAR_NOT_FOUND"
        return out
    out["_JOIN_SUBJID"] = normalize_join_value_series(out[subj_t])
    ref["_JOIN_SUBJID"] = normalize_join_value_series(ref[subj_r])
    ref_small = ref[["_JOIN_SUBJID", base_var]].copy()
    ref_small[base_var] = ref_small[base_var].replace("", pd.NA)
    ref_small = ref_small.groupby("_JOIN_SUBJID", as_index=False)[base_var].agg(first_non_missing)
    if raw_token in out.columns:
        out = out.drop(columns=[raw_token])
    out[raw_token] = out["_JOIN_SUBJID"].map(ref_small.set_index("_JOIN_SUBJID")[base_var])
    ref_obj.setdefault("MERGE_DEBUG", {}).update({"MERGE_METHOD": "SUBJECT_MAP", "OUTPUT_NON_NULL_N": int(out[raw_token].notna().sum()), "REF_NON_NULL_N": int(ref_small[base_var].notna().sum())})
    return out.drop(columns=["_JOIN_SUBJID"], errors="ignore")


# 用 SUBJID+VISIT 合併
def materialize_visit_ref(target_df, ref_df, raw_token, base_var, ref_obj):
    out, ref = normalize_upper_cols(target_df), normalize_upper_cols(ref_df)
    raw_token, base_var = str(raw_token).upper(), str(base_var).upper()
    subj_t, subj_r = get_subject_key(out), get_subject_key(ref)
    visit_t, visit_r = get_visit_key(out), get_visit_key(ref)
    if None in [subj_t, subj_r, visit_t, visit_r] or base_var not in ref.columns:
        out[raw_token] = pd.NA
        ref_obj.setdefault("MERGE_DEBUG", {})["ERROR"] = "VISIT_REF_KEY_OR_VAR_NOT_FOUND"
        return out
    out["_JOIN_SUBJID"] = normalize_join_value_series(out[subj_t]); ref["_JOIN_SUBJID"] = normalize_join_value_series(ref[subj_r])
    out["_JOIN_VISIT"] = normalize_join_value_series(out[visit_t]); ref["_JOIN_VISIT"] = normalize_join_value_series(ref[visit_r])
    ref_small = ref[["_JOIN_SUBJID", "_JOIN_VISIT", base_var]].copy()
    ref_small[base_var] = ref_small[base_var].replace("", pd.NA)
    ref_small = ref_small.groupby(["_JOIN_SUBJID", "_JOIN_VISIT"], as_index=False)[base_var].agg(first_non_missing)
    ref_small = ref_small.rename(columns={base_var: raw_token})
    if raw_token in out.columns:
        out = out.drop(columns=[raw_token])
    out = out.merge(ref_small, on=["_JOIN_SUBJID", "_JOIN_VISIT"], how="left")
    ref_obj.setdefault("MERGE_DEBUG", {}).update({"MERGE_METHOD": "VISIT_MERGE", "OUTPUT_NON_NULL_N": int(out[raw_token].notna().sum()), "REF_NON_NULL_N": int(ref_small[raw_token].notna().sum())})
    return out.drop(columns=["_JOIN_SUBJID", "_JOIN_VISIT"], errors="ignore")


# 先filter visit 再決定item level 判斷用 SUBJID/SUBJID+ITEM 合併 (例如: VISDAT_C1D1)
def materialize_fixed_visit_ref(target_df, ref_df, raw_token, base_var, visit_suffix, ref_obj):
    """
    規則：
      1. 先 filter ref_df 到指定 visit_suffix
      2. 如果 base_var 是 item-result variable，且 target/ref 有共同 item key：
            merge by SUBJID + item keys
         例如：
            HWORRES_S
            target/ref 都有 HWTEST
            => SUBJID + HWTEST
      3. 否則 fallback subject-level map
    """

    out = normalize_upper_cols(target_df)
    ref = normalize_upper_cols(ref_df)

    raw_token = str(raw_token).strip().upper()
    base_var = str(base_var).strip().upper()
    visit_suffix = str(visit_suffix).strip().upper()

    visit_key = get_visit_key(ref)

    if visit_key is None:
        out[raw_token] = pd.NA
        ref_obj.setdefault("MERGE_DEBUG", {})
        ref_obj["MERGE_DEBUG"]["ERROR"] = "REF_VISIT_KEY_NOT_FOUND_FOR_FIXED_VISIT"
        return out

    # =====================================================
    # 1. filter ref to fixed visit
    # =====================================================
    ref[visit_key] = normalize_join_value_series(ref[visit_key])
    ref = ref[ref[visit_key] == visit_suffix].copy()

    # =====================================================
    # 2. 判斷 base_var 是否是 item-result variable
    #    不 hardcode Height / Weight
    # =====================================================
    is_item_result_var = any(
        re.search(pat, base_var)
        for pat in ITEM_RESULT_SUFFIX_PATTERNS
    )

    # target / ref 共同 item keys
    common_item_keys = [
        c for c in out.columns
        if (
            c in ref.columns
            and re.search(ITEM_KEY_SUFFIX_PATTERN, str(c).strip().upper())
        )
    ]

    # =====================================================
    # 3. item-level fixed visit merge
    # =====================================================
    if is_item_result_var and common_item_keys:

        subj_t = get_subject_key(out)
        subj_r = get_subject_key(ref)

        if subj_t is None or subj_r is None:
            out[raw_token] = pd.NA
            ref_obj.setdefault("MERGE_DEBUG", {})
            ref_obj["MERGE_DEBUG"]["ERROR"] = "SUBJECT_KEY_NOT_FOUND_FOR_FIXED_VISIT_ITEM"
            return out

        if base_var not in ref.columns:
            out[raw_token] = pd.NA
            ref_obj.setdefault("MERGE_DEBUG", {})
            ref_obj["MERGE_DEBUG"]["ERROR"] = "BASE_VAR_NOT_FOUND_FOR_FIXED_VISIT_ITEM"
            return out

        # subject join
        out["_JOIN_SUBJID"] = normalize_join_value_series(out[subj_t])
        ref["_JOIN_SUBJID"] = normalize_join_value_series(ref[subj_r])

        # item joins
        join_cols = ["_JOIN_SUBJID"]

        for i, c in enumerate(common_item_keys):
            jc = f"_JOIN_ITEM_{i}"

            out[jc] = normalize_join_value_series(out[c])
            ref[jc] = normalize_join_value_series(ref[c])

            join_cols.append(jc)

        ref_small = ref[join_cols + [base_var]].copy()

        ref_small[base_var] = ref_small[base_var].replace({
            "": pd.NA,
            " ": pd.NA,
            "None": pd.NA,
            "NONE": pd.NA,
            "none": pd.NA,
            "nan": pd.NA,
            "NaN": pd.NA,
            "NAN": pd.NA,
            "<NA>": pd.NA,
            "NA": pd.NA,
            "N/A": pd.NA,
        })

        ref_small = (
            ref_small
            .groupby(join_cols, as_index=False, dropna=False)[base_var]
            .agg(first_non_missing)
        )

        ref_small = ref_small.rename(columns={base_var: raw_token})

        if raw_token in out.columns:
            out = out.drop(columns=[raw_token])

        out = out.merge(
            ref_small,
            on=join_cols,
            how="left"
        )

        ref_obj.setdefault("MERGE_DEBUG", {})
        ref_obj["MERGE_DEBUG"].update({
            "MERGE_METHOD": "FIXED_VISIT_ITEM_MERGE",
            "VISIT_SUFFIX": visit_suffix,
            "ITEM_JOIN_KEYS": common_item_keys,
            "LEFT_JOIN_ON": join_cols,
            "RIGHT_JOIN_ON": join_cols,
            "REF_ROWS_AFTER_VISIT_FILTER": len(ref),
            "REF_ROWS_AFTER_GROUP": len(ref_small),
            "OUTPUT_NON_NULL_N": int(out[raw_token].notna().sum()) if raw_token in out.columns else 0,
        })

        out = out.drop(columns=join_cols, errors="ignore")

        return out

    # =====================================================
    # 4. fallback: subject-level fixed visit
    # =====================================================
    return materialize_subject_ref(
        target_df=out,
        ref_df=ref,
        raw_token=raw_token,
        base_var=base_var,
        ref_obj=ref_obj
    )



# 用 X + VISIT_DAY 合併 (例如: VISDAT_CXD1)
def materialize_dynamic_cycle_ref(target_df, ref_df, ref_obj, raw_spec=None):
    out, ref = add_dynamic_visit_context_columns(target_df), add_dynamic_visit_context_columns(ref_df)
    raw_token = str(ref_obj.get("RAW_TOKEN", "")).strip().upper()
    base_var = str(ref_obj.get("BASE_VAR", "")).strip().upper()
    suffix = str(ref_obj.get("VISIT_SUFFIX", "")).strip().upper()
    subj_t, subj_r = get_subject_key(out), get_subject_key(ref)
    if subj_t is None or subj_r is None or base_var not in ref.columns:
        out[raw_token] = pd.NA
        ref_obj.setdefault("MERGE_DEBUG", {})["ERROR"] = "DYNAMIC_REF_KEY_OR_VAR_NOT_FOUND"
        return out
    m = re.fullmatch(r"CXD(\d+)", suffix)
    required_day = int(m.group(1)) if m else None
    out["_JOIN_SUBJID"] = normalize_join_value_series(out[subj_t]); ref["_JOIN_SUBJID"] = normalize_join_value_series(ref[subj_r])
    out["_JOIN_X"] = pd.to_numeric(out["X"], errors="coerce")
    out["_JOIN_DAY"] = required_day if required_day is not None else pd.to_numeric(out["VISIT_DAY"], errors="coerce")
    if required_day is not None:
        ref = ref[pd.to_numeric(ref["VISIT_DAY"], errors="coerce").eq(required_day)].copy()
    ref["_JOIN_X"] = pd.to_numeric(ref["X"], errors="coerce")
    ref["_JOIN_DAY"] = pd.to_numeric(ref["VISIT_DAY"], errors="coerce")
    ref_small = ref[["_JOIN_SUBJID", "_JOIN_X", "_JOIN_DAY", base_var]].copy()
    ref_small[base_var] = ref_small[base_var].replace("", pd.NA)
    ref_small = ref_small.groupby(["_JOIN_SUBJID", "_JOIN_X", "_JOIN_DAY"], as_index=False)[base_var].agg(first_non_missing)
    ref_small = ref_small.rename(columns={base_var: raw_token})
    if raw_token in out.columns:
        out = out.drop(columns=[raw_token])
    out = out.merge(ref_small, on=["_JOIN_SUBJID", "_JOIN_X", "_JOIN_DAY"], how="left")
    ref_obj.setdefault("MERGE_DEBUG", {}).update({"MERGE_METHOD": "DYNAMIC_CYCLE", "REQUIRED_DAY": required_day, "OUTPUT_NON_NULL_N": int(out[raw_token].notna().sum())})
    return out.drop(columns=["_JOIN_SUBJID", "_JOIN_X", "_JOIN_DAY"], errors="ignore")


# 統一入口：決定用哪個 merge
def materialize_one_ref(target_df, merged_dataset_map, ref_obj, raw_spec):
    out = normalize_upper_cols(target_df)
    raw_token = str(ref_obj.get("RAW_TOKEN", "")).strip().upper()
    base_var = str(ref_obj.get("BASE_VAR", "")).strip().upper()
    ref_stat = str(ref_obj.get("REF_STAT_DATASET", "")).strip().upper()
    join_type = str(ref_obj.get("JOIN_TYPE", "")).strip().upper()
    visit_suffix = ref_obj.get("VISIT_SUFFIX")
    ref_obj["MERGE_DEBUG"] = {"RAW_TOKEN": raw_token, "BASE_VAR": base_var, "REF_STAT_DATASET": ref_stat, "JOIN_TYPE": join_type}
    ref_df = merged_dataset_map.get(ref_stat)
    if ref_df is None or len(ref_df) == 0:
        out[raw_token] = pd.NA; ref_obj["MERGE_DEBUG"]["ERROR"] = "REF_NOT_FOUND"
        return out
    ref_df = attach_visit_order(ref_df, raw_spec)
    ref_df = normalize_upper_cols(ref_df)
    if join_type == "FIXED_VISIT":
        ref_obj["RESOLVED_JOIN_TYPE"] = "SUBJECT"
        return materialize_fixed_visit_ref(out, ref_df, raw_token, base_var, visit_suffix, ref_obj)
    if join_type == "DYNAMIC_CYCLE":
        ref_obj["RESOLVED_JOIN_TYPE"] = "DYNAMIC_CYCLE"
        return materialize_dynamic_cycle_ref(out, ref_df, ref_obj, raw_spec)
    # AUTO
    grain, dbg = resolve_auto_ref_grain(ref_df, base_var)
    ref_obj["RESOLVED_JOIN_TYPE"] = grain
    ref_obj["JOIN_GRAIN_DEBUG"] = dbg
    return materialize_visit_ref(out, ref_df, raw_token, base_var, ref_obj) if grain == "VISIT" else materialize_subject_ref(out, ref_df, raw_token, base_var, ref_obj)


# 收集一個 STAT_DATASET 所有 ref
def collect_refs_by_stat_dataset(dvp_review_df):
    out = {}
    if dvp_review_df is None or dvp_review_df.empty:
        return out
    for _, r in dvp_review_df.iterrows():
        stat = str(r.get("STAT_DATASET", "")).strip().upper()
        if not stat:
            continue
        out.setdefault(stat, [])
        for ref in r.get("REF_SPECS", []) or []:
            if isinstance(ref, dict):
                out[stat].append(copy.deepcopy(ref))
    for stat, refs in out.items():
        seen, uniq = set(), []
        for r in refs:
            key = (r.get("RAW_TOKEN"), r.get("BASE_VAR"), r.get("REF_STAT_DATASET"), r.get("JOIN_TYPE"), r.get("VISIT_SUFFIX"))
            if key not in seen:
                seen.add(key); uniq.append(r)
        out[stat] = uniq
    return out


# 對 target dataset 一次 merge 所有 ref
def enrich_stat_dataset_with_refs(target_df, refs, merged_dataset_map, raw_spec):
    out = attach_visit_order(target_df, raw_spec)
    out = normalize_upper_cols(out)
    out = add_dynamic_visit_context_columns(out)
    debug_rows = []
    for ref in refs or []:
        out = materialize_one_ref(out, merged_dataset_map, ref, raw_spec)
        out = add_dynamic_visit_context_columns(normalize_upper_cols(out))
        d = ref.get("MERGE_DEBUG", {}).copy()
        d.update({"RAW_TOKEN": ref.get("RAW_TOKEN"), "RESOLVED_JOIN_TYPE": ref.get("RESOLVED_JOIN_TYPE")})
        debug_rows.append(d)
    return out, pd.DataFrame(debug_rows)



# =========================================================
# Row ops
# =========================================================

# 決定 group key
def infer_partition_by(df, source_var=None):
    """
    規則：
    1. 一定包含 SUBJECT (SUBJID / USUBJID)
    2. 自動找與 source_var 對應的 item/test 欄位 (例如 VSTEST / LBTEST)
    """

    if df is None or df.empty:
        return []

    out = normalize_upper_cols(df)
    cols = [str(c).strip().upper() for c in out.columns]

    keys = []

    # 1. SUBJECT
    subj = get_subject_key(out)
    if subj:
        keys.append(subj)

    source_var = str(source_var or "").strip().upper()

    # 2. 找與 source_var 同 prefix 的 test/item (統一使用 ITEM_KEY_SUFFIX_PATTERN)
    matches = []

    for c in cols:
        
        if re.search(ITEM_KEY_SUFFIX_PATTERN, c):

            # 把後綴去掉，拿 prefix
            base = re.sub(
                ITEM_KEY_SUFFIX_PATTERN,
                "",
                c
            )

            base = str(base).strip().upper()

            if base and source_var.startswith(base):
                matches.append(c)


    # 3. 如果有 match → 用這組
    if matches:
        keys.extend(matches)

    else:
        # fallback（不靠 prefix）
        for c in cols:
            if re.search(ITEM_KEY_SUFFIX_PATTERN, c):
                keys.append(c)


    return list(dict.fromkeys(keys))




# 決定排序
def infer_order_by(df, source_var=None):
    """
    規則：
    1. 優先找與 source_var 同 prefix 的日期欄 (DAT/DATE/DTC)
    2. 再找其他 date-like
    3. 再用 VISIT_ORDER / VISITNUM
    """

    if df is None or df.empty:
        return []

    out = normalize_upper_cols(df)
    cols = [str(c).strip().upper() for c in out.columns]

    source_var = str(source_var or "").strip().upper()


    date_cols = [
        c for c in cols
        if _is_eval_date_like_column(c)
    ]

    order_cols = []

    # =====================================================
    # 1. 找 date-like 欄位
    # =====================================================
    date_cols = [
        c for c in cols
        if _is_eval_date_like_column(c)
    ]

    order_cols = []

    # =====================================================
    # 2. 優先找與 source_var 同 prefix 的 date
    #
    # 例：
    #   LBDAT -> base LB
    #   source_var LBCS startswith LB
    #
    #   CHEDAT -> base CHE
    #   source_var CHERES startswith CHE
    # =====================================================
    matched_date_cols = []

    for c in date_cols:

        base = re.sub(
            r"(DAT|DTC|DATE|DT)$",
            "",
            str(c).strip().upper()
        )

        if base and source_var.startswith(base):
            matched_date_cols.append(c)

    if matched_date_cols:
        order_cols.extend(matched_date_cols[:1])

    elif date_cols:
        order_cols.extend(date_cols[:1])

    # =====================================================
    # 3. tie-breaker
    # =====================================================
    for c in ["VISIT_ORDER", "VISITNUM"]:
        if c in cols and c not in order_cols:
            order_cols.append(c)

    return order_cols




# 判斷 row-op 的 source variable 來自哪個 STAT_DATASET。
def infer_row_op_source_stat_dataset(source_var, target_stat_dataset, raw_spec):
    """
      target=VS, source_var=VSCLISIG -> VS
      target=VS, source_var=EXDAT    -> EX

    如果找不到，fallback target_stat_dataset。
    """

    source_var = str(source_var).strip().upper()
    target_stat_dataset = str(target_stat_dataset).strip().upper()

    var_to_stat = build_var_to_stat_map(raw_spec)
    candidates = sorted(var_to_stat.get(source_var, []))

    if not candidates:
        return target_stat_dataset

    # 如果 target dataset 本身有這個 var，優先 target
    if target_stat_dataset in candidates:
        return target_stat_dataset

    # 否則取第一個 schema mapping dataset
    return candidates[0]



# 在單一 dataframe 內計算 row-op derived variables。
def compute_row_ops_within_df(df, row_ops, raw_spec=None):
    """
    適用：
      - same dataset row-op
      - cross dataset row-op 的 source dataset 先行計算

    重點：
      1. partition 使用 infer_partition_by()
      2. order 使用 infer_order_by()
      3. 日期優先排序
      4. 支援 SAS date number / string date
      5. shift 後用 _ORIG_ROW_ID 回填原始列
      6. 不 forward-fill，前一筆空值就保留空值

    例：
      LBCS_PREV：
        C2D1 的前一筆若 C1D15 LBCS 為空，
        則 C2D1 的 LBCS_PREV 應為空。
    """


    if df is None:
        return pd.DataFrame()

    out = normalize_upper_cols(df)

    if out.empty:
        return out

    if raw_spec is not None:
        out = attach_visit_order(out, raw_spec)

    out = normalize_upper_cols(out)

    # 保留原始 row id，排序後回填要用
    out["_ORIG_ROW_ID"] = range(len(out))

    for op in row_ops or []:

        if not isinstance(op, dict):
            continue

        source_var = str(op.get("SOURCE_VAR", "")).strip().upper()
        derived_var = str(op.get("DERIVED_VAR", "")).strip().upper()
        op_type = str(op.get("OP", "")).strip().upper()

        if not source_var or not derived_var:
            continue

        if source_var not in out.columns:
            out[derived_var] = pd.NA
            op["ROW_OP_DEBUG"] = {
                "SOURCE_VAR": source_var,
                "DERIVED_VAR": derived_var,
                "OP": op_type,
                "ERROR": "SOURCE_VAR_NOT_FOUND",
            }
            continue

        # =====================================================
        # partition / order
        # =====================================================
        partition_by = infer_partition_by(
            out,
            source_var=source_var
        )

        order_by = infer_order_by(
            out,
            source_var=source_var
        )

        partition_by = [
            c for c in partition_by
            if c in out.columns
        ]

        order_by = [
            c for c in order_by
            if c in out.columns
        ]

        tmp = out.copy()

        # =====================================================
        # 建 normalized partition columns
        # 避免 LBTEST 前後空白 / 大小寫造成 group 錯
        # =====================================================
        part_tmp_cols = []

        for i, c in enumerate(partition_by):

            pc = f"_ROWOP_PART_{i}"

            tmp[pc] = (
                tmp[c]
                .astype(str)
                .str.strip()
                .str.upper()
                .replace({
                    "": pd.NA,
                    "NAN": pd.NA,
                    "NONE": pd.NA,
                    "<NA>": pd.NA,
                })
            )

            part_tmp_cols.append(pc)

        # =====================================================
        # 建排序欄位
        # =====================================================
        order_tmp_cols = []

        for i, c in enumerate(order_by):

            oc = f"_ROWOP_ORDER_{i}"

            if _is_eval_date_like_column(c):

                # numeric date, SAS origin
                num = pd.to_numeric(
                    tmp[c],
                    errors="coerce"
                )

                dt_from_num = pd.to_datetime(
                    num,
                    errors="coerce",
                    unit="D",
                    origin="1960-01-01"
                )

                # text date
                dt_from_text = pd.to_datetime(
                    tmp[c].astype(str).str.replace("/", "-"),
                    errors="coerce"
                )

                tmp[oc] = dt_from_num.where(
                    num.notna(),
                    dt_from_text
                )

            else:
                num = pd.to_numeric(
                    tmp[c],
                    errors="coerce"
                )

                tmp[oc] = num.where(
                    num.notna(),
                    tmp[c].astype(str).str.strip()
                )

            order_tmp_cols.append(oc)

        # 如果沒有 order_by，至少用 VISITNUM / VISIT_ORDER fallback
        if not order_tmp_cols:

            for fallback_col in ["VISIT_ORDER", "VISITNUM"]:

                if fallback_col in tmp.columns:

                    oc = f"_ROWOP_ORDER_{len(order_tmp_cols)}"

                    tmp[oc] = pd.to_numeric(
                        tmp[fallback_col],
                        errors="coerce"
                    )

                    order_tmp_cols.append(oc)
                    break

        # deterministic tie-breaker
        sort_cols = part_tmp_cols + order_tmp_cols + ["_ORIG_ROW_ID"]

        tmp = tmp.sort_values(
            sort_cols,
            na_position="last"
        ).copy()

        # =====================================================
        # 重要：
        # 這裡不要 first_non_missing / ffill
        # PREV 就是上一列原始值，上一列空就應該是空
        # =====================================================
        source_series = tmp[source_var].replace({
            "": pd.NA,
            " ": pd.NA,
            "None": pd.NA,
            "NONE": pd.NA,
            "none": pd.NA,
            "nan": pd.NA,
            "NaN": pd.NA,
            "NAN": pd.NA,
            "<NA>": pd.NA,
            "NA": pd.NA,
            "N/A": pd.NA,
        })

        if part_tmp_cols:
            grp = source_series.groupby(
                [tmp[c] for c in part_tmp_cols],
                dropna=False
            )
        else:
            grp = None

        if op_type in ["PREV", "PREVIOUS"]:

            tmp[derived_var] = (
                grp.shift(1)
                if grp is not None
                else source_series.shift(1)
            )

        elif op_type == "NEXT":

            tmp[derived_var] = (
                grp.shift(-1)
                if grp is not None
                else source_series.shift(-1)
            )

        elif op_type == "FIRST":

            tmp[derived_var] = (
                grp.transform("first")
                if grp is not None
                else source_series.iloc[0]
            )

        elif op_type == "LAST":

            tmp[derived_var] = (
                grp.transform("last")
                if grp is not None
                else source_series.iloc[-1]
            )

        else:

            tmp[derived_var] = pd.NA

            op["ROW_OP_DEBUG"] = {
                "SOURCE_VAR": source_var,
                "DERIVED_VAR": derived_var,
                "OP": op_type,
                "ERROR": f"UNKNOWN_ROW_OP: {op_type}",
            }

        # =====================================================
        # 回填原始列
        # =====================================================
        map_series = tmp.set_index("_ORIG_ROW_ID")[derived_var]

        out[derived_var] = out["_ORIG_ROW_ID"].map(map_series)

        # =====================================================
        # Debug info
        # =====================================================
        op["ROW_OP_DEBUG"] = {
            "SOURCE_VAR": source_var,
            "DERIVED_VAR": derived_var,
            "OP": op_type,
            "PARTITION_BY": partition_by,
            "ORDER_BY": order_by,
            "SORT_COLS": sort_cols,
            "OUTPUT_NON_NULL_N": (
                int(out[derived_var].notna().sum())
                if derived_var in out.columns
                else 0
            ),
        }

        # 清暫存欄位
        tmp = tmp.drop(
            columns=part_tmp_cols + order_tmp_cols,
            errors="ignore"
        )

    out = out.drop(
        columns=["_ORIG_ROW_ID"],
        errors="ignore"
    )

    return out




# 將 cross-dataset row-op 結果 merge 回 target_df。
def materialize_cross_dataset_row_op(target_df, source_df, row_op, raw_spec=None):
    """
    例：
      target = VS
      source = EX
      row_op = EXDAT_FIRST

    步驟：
      1. 在 EX 裡算 EXDAT_FIRST
      2. 根據 op 類型決定 merge grain
         FIRST/LAST -> SUBJECT
         PREV/NEXT  -> VISIT if possible, else SUBJECT
      3. merge 回 VS
    """

    out = normalize_upper_cols(target_df)
    src = normalize_upper_cols(source_df)

    source_var = str(row_op.get("SOURCE_VAR", "")).strip().upper()
    derived_var = str(row_op.get("DERIVED_VAR", "")).strip().upper()
    op_type = str(row_op.get("OP", "")).strip().upper()

    if not source_var or not derived_var:
        return out

    # 先在 source dataset 算 derived_var
    src = compute_row_ops_within_df(
        src,
        [row_op],
        raw_spec=raw_spec
    )

    subj_t = get_subject_key(out)
    subj_s = get_subject_key(src)

    if subj_t is None or subj_s is None:
        out[derived_var] = pd.NA
        row_op.setdefault("ROW_OP_DEBUG", {})
        row_op["ROW_OP_DEBUG"]["ERROR"] = "SUBJECT_KEY_NOT_FOUND_FOR_CROSS_ROW_OP"
        return out

    out["_JOIN_SUBJID"] = normalize_join_value_series(out[subj_t])
    src["_JOIN_SUBJID"] = normalize_join_value_series(src[subj_s])

    # FIRST / LAST 通常是 subject-level result
    if op_type in ["FIRST", "LAST"]:

        ref_small = src[["_JOIN_SUBJID", derived_var]].copy()
        ref_small[derived_var] = ref_small[derived_var].replace("", pd.NA)

        ref_small = (
            ref_small
            .groupby("_JOIN_SUBJID", as_index=False)[derived_var]
            .agg(first_non_missing)
        )

        if derived_var in out.columns:
            out = out.drop(columns=[derived_var])

        out[derived_var] = out["_JOIN_SUBJID"].map(
            ref_small.set_index("_JOIN_SUBJID")[derived_var]
        )

        row_op.setdefault("ROW_OP_DEBUG", {})
        row_op["ROW_OP_DEBUG"].update({
            "MERGE_METHOD": "CROSS_ROW_OP_SUBJECT_MAP",
            "SOURCE_STAT_DATASET": row_op.get("SOURCE_STAT_DATASET"),
            "TARGET_STAT_DATASET": row_op.get("TARGET_STAT_DATASET"),
            "OUTPUT_NON_NULL_N": int(out[derived_var].notna().sum())
        })

        out = out.drop(columns=["_JOIN_SUBJID"], errors="ignore")
        return out

    # PREV/NEXT 對 cross dataset 比較不一定有語意，
    # 如果 source/target 都有 visit key，就用 SUBJECT+VISIT merge。
    visit_t = get_visit_key(out)
    visit_s = get_visit_key(src)

    if visit_t is not None and visit_s is not None:

        out["_JOIN_VISIT"] = normalize_join_value_series(out[visit_t])
        src["_JOIN_VISIT"] = normalize_join_value_series(src[visit_s])

        ref_small = src[["_JOIN_SUBJID", "_JOIN_VISIT", derived_var]].copy()
        ref_small[derived_var] = ref_small[derived_var].replace("", pd.NA)

        ref_small = (
            ref_small
            .groupby(["_JOIN_SUBJID", "_JOIN_VISIT"], as_index=False)[derived_var]
            .agg(first_non_missing)
        )

        if derived_var in out.columns:
            out = out.drop(columns=[derived_var])

        out = out.merge(
            ref_small,
            on=["_JOIN_SUBJID", "_JOIN_VISIT"],
            how="left"
        )

        row_op.setdefault("ROW_OP_DEBUG", {})
        row_op["ROW_OP_DEBUG"].update({
            "MERGE_METHOD": "CROSS_ROW_OP_VISIT_MERGE",
            "OUTPUT_NON_NULL_N": int(out[derived_var].notna().sum()) if derived_var in out.columns else 0
        })

        out = out.drop(columns=["_JOIN_SUBJID", "_JOIN_VISIT"], errors="ignore")
        return out

    # fallback subject-level
    ref_small = src[["_JOIN_SUBJID", derived_var]].copy()

    ref_small = (
        ref_small
        .groupby("_JOIN_SUBJID", as_index=False)[derived_var]
        .agg(first_non_missing)
    )

    if derived_var in out.columns:
        out = out.drop(columns=[derived_var])

    out[derived_var] = out["_JOIN_SUBJID"].map(
        ref_small.set_index("_JOIN_SUBJID")[derived_var]
    )

    out = out.drop(columns=["_JOIN_SUBJID"], errors="ignore")

    row_op.setdefault("ROW_OP_DEBUG", {})
    row_op["ROW_OP_DEBUG"].update({
        "MERGE_METHOD": "CROSS_ROW_OP_SUBJECT_FALLBACK",
        "OUTPUT_NON_NULL_N": int(out[derived_var].notna().sum())
    })

    return out



# 產生 PREV/NEXT/FIRST/LAST
def apply_row_ops(target_df, row_ops, merged_dataset_map, target_stat_dataset, raw_spec=None):
    """
    支援：
      1. same dataset row-op
      2. cross dataset row-op

    例：
      target=VS
      VSCLISIG_PREV -> same dataset
      EXDAT_FIRST   -> source=EX, merge 回 VS
    """

    if target_df is None:
        return pd.DataFrame()

    out = normalize_upper_cols(target_df)

    target_stat_dataset = str(target_stat_dataset).strip().upper()

    same_ops = []
    cross_ops = []

    for op in row_ops or []:

        source_stat = str(
            op.get("SOURCE_STAT_DATASET")
            or target_stat_dataset
        ).strip().upper()

        if source_stat == target_stat_dataset:
            same_ops.append(op)
        else:
            cross_ops.append(op)

    # 1. same dataset row ops
    if same_ops:
        out = compute_row_ops_within_df(
            out,
            same_ops,
            raw_spec=raw_spec
        )

    # 2. cross dataset row ops
    for op in cross_ops:

        source_stat = str(op.get("SOURCE_STAT_DATASET", "")).strip().upper()

        source_df = merged_dataset_map.get(source_stat)

        if source_df is None or len(source_df) == 0:
            derived_var = str(op.get("DERIVED_VAR", "")).strip().upper()
            if derived_var:
                out[derived_var] = pd.NA

            op.setdefault("ROW_OP_DEBUG", {})
            op["ROW_OP_DEBUG"]["ERROR"] = "SOURCE_DATASET_NOT_FOUND"
            continue

        out = materialize_cross_dataset_row_op(
            target_df=out,
            source_df=source_df,
            row_op=op,
            raw_spec=raw_spec
        )

    return out





# =========================================================
# Eval helpers
# =========================================================

def _series_to_sas_date_num(s):
    if s is None:
        return pd.Series(dtype="float64")
    num = pd.to_numeric(s, errors="coerce")
    non_empty = s.dropna()
    non_empty = non_empty[non_empty.astype(str).str.strip() != ""]
    if len(non_empty) > 0 and num.notna().sum() == len(non_empty):
        return num
    dt = pd.to_datetime(s, errors="coerce")
    return (dt - pd.Timestamp("1960-01-01")).dt.days


def _is_eval_date_like_column(col):
    cu = str(col).upper()
    if cu in ["X", "VISIT_DAY", "VISIT_ORDER", "VISITNUM"] or cu.endswith("_N"):
        return False
    return bool(cu.endswith("DAT") or cu.endswith("DTC") or cu.endswith("DATE") or cu.endswith("DT") or "STDAT" in cu or "ENDAT" in cu or re.search(r"(DAT|DTC|DATE|DT)_", cu))


# 把 DATE → numeric 把數值轉型
def normalize_eval_df_for_eval(df):
    out = normalize_upper_cols(df)
    for c in out.columns:
        cu, s = str(c).upper(), out[c]
        if cu in ["X", "VISIT_DAY", "VISIT_ORDER", "VISITNUM"] or cu.endswith("_N") or cu.endswith("NUM"):
            out[c] = pd.to_numeric(s, errors="coerce"); continue
        if _is_eval_date_like_column(c):
            out[c] = _series_to_sas_date_num(s); continue
        non_empty = s.dropna(); non_empty = non_empty[non_empty.astype(str).str.strip() != ""]
        if len(non_empty) > 0:
            num = pd.to_numeric(s, errors="coerce")
            if num.notna().sum() == len(non_empty): out[c] = num
    return out


# 轉邏輯字元
def normalize_eval_condition_text(condition):
    s = clean_condition_text(condition)
    s = s.replace("[", "(").replace("]", ")")
    s = s.replace("^", "**")
    op_map = [(r"\bgte\b", ">="), (r"\blte\b", "<="), (r"\bge\b", ">="), (r"\ble\b", "<="), (r"\bgt\b", ">"), (r"\blt\b", "<"), (r"\beq\b", "=="), (r"\bne\b", "!=")]
    for pat, repl in op_map:
        s = re.sub(pat, repl, s, flags=re.I)
    s = re.sub(r"(?<![<>=!])=(?![=])", "==", s)

    # =====================================================
    # 修正常見 DVP 寫法：
    #   AECTCAE == "Grade 3" or "Grade 4" or "Grade 5"
    # 轉成：
    #   (AECTCAE == "Grade 3" or AECTCAE == "Grade 4" or AECTCAE == "Grade 5")
    # =====================================================
    def expand_or_values(m):
        var = m.group(1)
        first_val = m.group(2)
        rest = m.group(3)
        vals = [first_val]
        vals.extend(re.findall(r'\bor\s+("[^"]*"|\'[^\']*\')',rest,flags=re.IGNORECASE))
        parts = [f"{var} == {v}"for v in vals]

        return "(" + " or ".join(parts) + ")"

    s = re.sub(
        r'\b([A-Za-z_][A-Za-z0-9_]*)\s*==\s*("[^"]*"|\'[^\']*\')((?:\s+or\s+("[^"]*"|\'[^\']*\'))+)',
        expand_or_values,
        s,
        flags=re.IGNORECASE
    )

    return s


# 轉極值
def convert_rowwise_max_min_functions(s, col_set):
    def repl(m):
        func, a, b = m.group(1).lower(), m.group(2).upper(), m.group(3).upper()
        if a not in col_set or b not in col_set: return m.group(0)
        method = "max" if func == "max" else "min"
        return f'pd.concat([df["{a}"], df["{b}"]], axis=1).{method}(axis=1)'
    return re.sub(r"\b(max|min)\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)", repl, str(s), flags=re.I)


# 自動補 notna()，避免missing的大小比較
def add_non_missing_guard_for_comparisons(expr):
    s = str(expr)

    df_col_pat = r'df\["[A-Za-z_][A-Za-z0-9_]*"\]'

    # 抓一段 comparison：
    # left comparison right
    # right 直到遇到 boolean operator & / | 或字串結尾
    comp_pat = re.compile(
        rf'({df_col_pat})\s*(==|!=|>=|<=|>|<)\s*(.*?)(?=\s+[\&\|]\s+|$)'
    )

    def collect_notna_guards(text):
        guards = []

        for c in re.findall(df_col_pat, str(text)):
            guards.append(f"{c}.notna()")

        return list(dict.fromkeys(guards))

    def repl(m):

        left = m.group(1)
        op = m.group(2)
        right = m.group(3).strip()

        # 避免重複包已經被處理過的 comparison
        if right == "":
            return m.group(0)


        # =====================================================
        # 如果是複雜公式，不要包 guard，避免切壞
        # 例：
        #   year(LBDAT) - BRTHDTC
        #   TORRES / ((RRORRES*0.001) ** (1/3))
        # =====================================================
        complex_rhs = bool(
            re.search(r"[\+\-\*/]", right)
            or "**" in right
            or "pd.to_datetime" in right
            or ".dt.year" in right
        )

        if complex_rhs:
            return f"({left} {op} {right})"

        guards = []
        guards.extend(collect_notna_guards(left))
        guards.extend(collect_notna_guards(right))
        guards = list(dict.fromkeys(guards))

        comparison = f"({left} {op} {right})"

        if guards:
            return "(" + " & ".join(guards) + " & " + comparison + ")"

        return comparison

    return comp_pat.sub(repl, s)




# 如果欄位不存在 → 補 NA
def ensure_eval_columns(eval_df, condition):
    out = normalize_upper_cols(eval_df)
    for tok in find_identifiers_in_condition(condition):
        meta = normalize_token_meta(tok)
        raw = meta["RAW_TOKEN"]
        if raw not in out.columns and raw not in RESERVED_TOKENS:
            out[raw] = pd.NA
    return out


# 將 DVP condition 轉成 pandas 可 eval 的 Python expression。
def condition_to_python_expr(condition, columns):

    s = normalize_eval_condition_text(condition)

    if not s:
        return ""

    col_set = {str(c).strip().upper()for c in columns}

    # Protect quoted strings
    string_map = {}

    def protect_string(m):
        key = f"__STR_{len(string_map)}__"
        string_map[key] = m.group(0)
        return key

    s = re.sub(r'"[^"]*"|\'[^\']*\'', protect_string, s)

    # max(A,B) / min(A,B) -> row-wise max/min
    s = convert_rowwise_max_min_functions(s, col_set)


    # year(VAR) -> pd.to_datetime(df["VAR"]).dt.year
    def convert_year_func(m):
        var = str(m.group(1)).strip().upper()

        if var in col_set:
            return (
                "("
                f"pd.to_datetime("
                f"pd.to_numeric(df[\"{var}\"], errors=\"coerce\"), "
                f"errors=\"coerce\", "
                f"unit=\"D\", "
                f"origin=\"1960-01-01\""
                f").dt.year"
                f".where("
                f"pd.to_numeric(df[\"{var}\"], errors=\"coerce\").notna(), "
                f"pd.to_datetime(df[\"{var}\"], errors=\"coerce\").dt.year"
                f")"
                ")"
            )

        return m.group(0)

    s = re.sub(
        r"\byear\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)",
        convert_year_func,
        s,
        flags=re.IGNORECASE
    )


    # X parameter
    if "X" in col_set:
        s = re.sub(r"\bX\b", 'df["X"]', s)


    # Replace column names with df["COL"]
    for col in sorted(col_set, key=len, reverse=True):

        if col == "X":
            continue

        pattern = rf'(?<!["\'])\b{re.escape(col)}\b(?!["\'])'

        s = re.sub(
            pattern,
            f'df["{col}"]',
            s,
            flags=re.IGNORECASE
        )

    # Restore quoted strings
    for key, val in string_map.items():
        s = s.replace(key, val)

    # Boolean operators
    s = re.sub(r"\band\b", "&", s, flags=re.IGNORECASE)
    s = re.sub(r"\bor\b", "|", s, flags=re.IGNORECASE)


    # Add notna guards around comparisons
    s = add_non_missing_guard_for_comparisons(s)


    # Final wrap
    s = f"({s})"

    # balance parentheses if needed
    if s.count("(") > s.count(")"):
        s += ")" * (s.count("(") - s.count(")"))

    return re.sub(r"\s+", " ", s).strip()




def evaluate_rule_condition(eval_df, condition):
    eval_df = ensure_eval_columns(eval_df, condition)
    eval_df = normalize_eval_df_for_eval(eval_df)
    expr = condition_to_python_expr(condition, eval_df.columns)
    if not expr:
        return expr, pd.Series(False, index=eval_df.index), eval_df
    mask = eval(expr, {"__builtins__": {}}, {"df": eval_df, "pd": pd, "np": np, "re": re})
    if not isinstance(mask, pd.Series):
        mask = pd.Series(mask, index=eval_df.index)
    mask = mask.reindex(eval_df.index).fillna(False).astype(bool)
    return expr, mask, eval_df






# =========================================================
# Step 2 entry
# =========================================================
def run_dvp_data_check(dvp_review_df, merged_dataset_map, raw_spec):
    """
    流程：
      1. 將 Step1 不可執行 / 有問題的 rules 先列為 SKIP_STEP2
      2. 只對 READY_FOR_STEP2=True 的 rules 執行
      3. 依 STAT_DATASET group
      4. 每個 STAT_DATASET 一次 enrich refs
      5. 每個 STAT_DATASET 一次 apply row ops
      6. 每條 rule/block 再依 visit scope filter
      7. evaluate condition
    """

    results = []
    fail_detail_map = {}
    enriched_dataset_map = {}
    merge_debug_rows = []

    if dvp_review_df is None or dvp_review_df.empty:
        return (
            pd.DataFrame(),
            {},
            {},
            pd.DataFrame()
        )

    all_df = dvp_review_df.copy()

    # 確保 READY_FOR_STEP2 存在
    if "READY_FOR_STEP2" not in all_df.columns:
        all_df["READY_FOR_STEP2"] = False

    skipped_df = all_df[
        ~all_df["READY_FOR_STEP2"].astype(bool)
    ].copy()

    ready_df = all_df[
        all_df["READY_FOR_STEP2"].astype(bool)
    ].copy()

    # =====================================================
    # 1. 先把不可執行的 rules 放進 result
    # =====================================================
    for idx, row in skipped_df.iterrows():

        result = {
            "IDX": idx,
            "RULE": row.get("RULE"),
            "FORM": row.get("FORM"),
            "PD": row.get("PD"),
            "STAT_DATASET": row.get("STAT_DATASET"),
            "BLOCK_ID": row.get("BLOCK_ID"),
            "EFFECTIVE_VISIT": row.get("EFFECTIVE_VISIT"),
            "RULE_TYPE": row.get("RULE_TYPE"),
            "EXECUTION_CLASS": row.get("EXECUTION_CLASS"),
            "STATUS": "SKIP_STEP2",
            "FAIL_COUNT": 0,
            "TARGET_VARS": row.get("TARGET_VARS", []),
            "REF_VARS": row.get("REF_VARS", []),
            "REF_DATASETS": row.get("REF_DATASETS", []),
            "MESSAGE": row.get("MESSAGE", ""),
            "CONDITION_BLOCK": row.get("CONDITION_BLOCK", ""),
            "PYTHON_EXPR": "",
            "VISIT_SCOPE": row.get("EFFECTIVE_VISIT_SCOPE", []),
            "BLOCKING_REASON": row.get("BLOCKING_REASON", ""),
        }

        results.append(result)

        fail_detail_map[idx] = {
            **result,
            "FAIL_DF": pd.DataFrame(),
            "EVAL_DF": pd.DataFrame(),
        }

    # 如果沒有可執行 rule，就直接回傳 skipped result
    if ready_df.empty:
        result_df = pd.DataFrame(results)
        return (
            result_df,
            fail_detail_map,
            enriched_dataset_map,
            pd.DataFrame()
        )

    # =====================================================
    # 2. Collect refs by STAT_DATASET
    # =====================================================
    ref_map = collect_refs_by_stat_dataset(ready_df)

    # =====================================================
    # 3. 每個 STAT_DATASET 一次 enrich
    # =====================================================
    for stat, g in ready_df.groupby("STAT_DATASET"):

        stat = str(stat).strip().upper()

        target_df = merged_dataset_map.get(stat)

        if target_df is None or len(target_df) == 0:

            for idx, row in g.iterrows():

                result = {
                    "IDX": idx,
                    "RULE": row.get("RULE"),
                    "STAT_DATASET": stat,
                    "BLOCK_ID": row.get("BLOCK_ID"),
                    "EFFECTIVE_VISIT": row.get("EFFECTIVE_VISIT"),
                    "RULE_TYPE": row.get("RULE_TYPE"),
                    "EXECUTION_CLASS": row.get("EXECUTION_CLASS", "EXECUTABLE"),
                    "STATUS": "TARGET_DATASET_NOT_FOUND",
                    "FAIL_COUNT": 0,
                    "TARGET_VARS": row.get("TARGET_VARS", []),
                    "REF_VARS": row.get("REF_VARS", []),
                    "REF_DATASETS": row.get("REF_DATASETS", []),
                    "MESSAGE": row.get("MESSAGE", ""),
                    "CONDITION_BLOCK": row.get("CONDITION_BLOCK", ""),
                    "PYTHON_EXPR": "",
                    "VISIT_SCOPE": row.get("EFFECTIVE_VISIT_SCOPE", []),
                    "BLOCKING_REASON": "Target STAT_DATASET not found in uploaded data.",
                }

                results.append(result)

                fail_detail_map[idx] = {
                    **result,
                    "FAIL_DF": pd.DataFrame(),
                    "EVAL_DF": pd.DataFrame(),
                }

            continue

        # -------------------------------------------------
        # 3.1 Enrich refs once by STAT_DATASET
        # -------------------------------------------------
        refs = ref_map.get(stat, [])

        enriched_df, merge_debug_df = enrich_stat_dataset_with_refs(
            target_df=target_df,
            refs=refs,
            merged_dataset_map=merged_dataset_map,
            raw_spec=raw_spec
        )

        if merge_debug_df is not None and not merge_debug_df.empty:
            merge_debug_df = merge_debug_df.copy()
            merge_debug_df.insert(0, "STAT_DATASET", stat)
            merge_debug_rows.append(merge_debug_df)

        # -------------------------------------------------
        # 3.2 Apply row ops once by STAT_DATASET
        # -------------------------------------------------
        all_row_ops = []

        for _, rr in g.iterrows():
            ops = rr.get("ROW_OPS", []) or []

            if isinstance(ops, list):
                all_row_ops.extend(ops)

        # dedup row ops
        seen_ops = set()
        uniq_ops = []

        for op in all_row_ops:

            if not isinstance(op, dict):
                continue

            key = (
                op.get("SOURCE_VAR"),
                op.get("SOURCE_STAT_DATASET"),
                op.get("TARGET_STAT_DATASET"),
                op.get("DERIVED_VAR"),
                op.get("OP"),
            )

            if key not in seen_ops:
                seen_ops.add(key)
                uniq_ops.append(op)

        if uniq_ops:

            if "apply_row_ops" in globals():
                enriched_df = apply_row_ops(
                    target_df=enriched_df,
                    row_ops=uniq_ops,
                    merged_dataset_map=merged_dataset_map,
                    target_stat_dataset=stat,
                    raw_spec=raw_spec
                )

        enriched_dataset_map[stat] = enriched_df

        # =====================================================
        # 4. Per rule / block evaluate
        # =====================================================
        for idx, row in g.iterrows():

            rr = row.to_dict()

            eval_df = filter_eval_df_by_visit_scope(
                enriched_df,
                rr
            )

            if eval_df is None or eval_df.empty:

                result = {
                    "IDX": idx,
                    "RULE": rr.get("RULE"),
                    "FORM": rr.get("FORM"),
                    "PD": rr.get("PD"),
                    "STAT_DATASET": stat,
                    "BLOCK_ID": rr.get("BLOCK_ID"),
                    "EFFECTIVE_VISIT": rr.get("EFFECTIVE_VISIT"),
                    "RULE_TYPE": rr.get("RULE_TYPE"),
                    "EXECUTION_CLASS": rr.get("EXECUTION_CLASS", "EXECUTABLE"),
                    "STATUS": "OK",
                    "FAIL_COUNT": 0,
                    "TARGET_VARS": rr.get("TARGET_VARS", []),
                    "REF_VARS": rr.get("REF_VARS", []),
                    "REF_DATASETS": rr.get("REF_DATASETS", []),
                    "MESSAGE": rr.get("MESSAGE", ""),
                    "CONDITION_BLOCK": rr.get("CONDITION_BLOCK", ""),
                    "PYTHON_EXPR": "",
                    "VISIT_SCOPE": rr.get("EFFECTIVE_VISIT_SCOPE", []),
                    "BLOCKING_REASON": "",
                }

                results.append(result)

                fail_detail_map[idx] = {
                    **result,
                    "FAIL_DF": pd.DataFrame(),
                    "EVAL_DF": pd.DataFrame(),
                }

                continue

            try:
                eval_df_display = eval_df.copy()

                expr, mask, eval_df2 = evaluate_rule_condition(
                    eval_df,
                    rr.get("CONDITION_BLOCK", "")
                )

                mask = mask.reindex(eval_df_display.index).fillna(False).astype(bool)

                raw_fail_df = eval_df_display.loc[mask].copy()

                status = "OK"

                # =====================================================
                # Decide fail grain: item-level vs non-item-level
                # =====================================================
                target_vars = rr.get("TARGET_VARS", []) or []
                ref_vars = rr.get("REF_VARS", []) or []
                condition_tokens = rr.get("RULE_VAR_TOKENS", []) or []

                report_vars = [
                    str(x).strip().upper()
                    for x in list(target_vars) + list(ref_vars)
                    if str(x).strip()
                ]


                # Dataset item keys (不 hardcode domain prefix，直接看欄位 suffix)
                item_cols = [
                    c for c in raw_fail_df.columns
                    if re.search(
                        ITEM_KEY_SUFFIX_PATTERN,
                        str(c).strip().upper()
                    )
                ]

                has_item_keys_in_dataset = len(item_cols) > 0


                # condition 是否明確使用 item/test 欄位
                has_explicit_item_condition = any(
                    re.search(
                        ITEM_KEY_SUFFIX_PATTERN,
                        str(t).strip().upper()
                    )
                    for t in condition_tokens
                )


                # condition 是否使用 item result 欄位 (例如: VSSTAT、VSCLISIG)，常數區ITEM_RESULT_SUFFIX_PATTERNS規範
                uses_item_result_var = any(
                    any(
                        re.search(pat, str(v).strip().upper())
                        for pat in ITEM_RESULT_SUFFIX_PATTERNS
                    )
                    for v in report_vars
                )


                # Final item-level decision
                is_item_level = (
                    has_item_keys_in_dataset
                    and (
                        has_explicit_item_condition
                        or uses_item_result_var
                    )
                )

                # =====================================================
                # Deduplicate fail rows by rule grain
                # =====================================================
                subj_col = get_subject_key(raw_fail_df)
                
                dedup_cols = []

                if subj_col:
                    dedup_cols.append(subj_col)

                if "VISIT" in raw_fail_df.columns:
                    dedup_cols.append("VISIT")

                if "VISIT_DVP" in raw_fail_df.columns:
                    dedup_cols.append("VISIT_DVP")


                # -------------------------------------------------
                # item-level: subject + visit + item/test
                # -------------------------------------------------
                if is_item_level:
                    dedup_cols.extend(item_cols)

                dedup_cols = [
                    c for c in list(dict.fromkeys(dedup_cols))
                    if c in raw_fail_df.columns
                ]


                if dedup_cols:
                    fail_df = raw_fail_df.drop_duplicates(
                        subset=dedup_cols,
                        keep="first"
                    ).copy()


                else:
                    fail_df = raw_fail_df.copy()


            except Exception as e:
                expr = ""
                raw_fail_df = pd.DataFrame()
                fail_df = pd.DataFrame()
                eval_df_display = eval_df.copy()
                eval_df2 = eval_df
                status = f"EVAL_ERROR: {e}"

                is_item_level = False
                dedup_cols = []


            result = {
                "IDX": idx,
                "RULE": rr.get("RULE"),
                "FORM": rr.get("FORM"),
                "PD": rr.get("PD"),
                "STAT_DATASET": stat,
                "BLOCK_ID": rr.get("BLOCK_ID"),
                "EFFECTIVE_VISIT": rr.get("EFFECTIVE_VISIT"),
                "RULE_TYPE": rr.get("RULE_TYPE"),
                "EXECUTION_CLASS": rr.get("EXECUTION_CLASS", "EXECUTABLE"),
                "STATUS": status,
                "FAIL_COUNT": len(fail_df) if isinstance(fail_df, pd.DataFrame) else 0,
                "RAW_FAIL_COUNT": len(raw_fail_df) if isinstance(raw_fail_df, pd.DataFrame) else 0,
                "IS_ITEM_LEVEL": bool(is_item_level),
                "FAIL_GRAIN_KEYS": dedup_cols,
                "TARGET_VARS": rr.get("TARGET_VARS", []),
                "REF_VARS": rr.get("REF_VARS", []),
                "REF_DATASETS": rr.get("REF_DATASETS", []),
                "MESSAGE": rr.get("MESSAGE", ""),
                "CONDITION_BLOCK": rr.get("CONDITION_BLOCK", ""),
                "PYTHON_EXPR": expr,
                "VISIT_SCOPE": rr.get("EFFECTIVE_VISIT_SCOPE", []),
                "BLOCKING_REASON": rr.get("BLOCKING_REASON", ""),
            }

            results.append(result)

            fail_detail_map[idx] = {
                **result,
                "FAIL_DF": fail_df,
                "RAW_FAIL_DF": raw_fail_df,
                "EVAL_DF": eval_df_display,
                "EVAL_DF_FOR_EVAL": eval_df2,
                "IS_ITEM_LEVEL": bool(is_item_level),
                "FAIL_GRAIN_KEYS": dedup_cols,
            }

    result_df = pd.DataFrame(results)

    if merge_debug_rows:
        merge_debug_all_df = pd.concat(
            merge_debug_rows,
            ignore_index=True
        )
    else:
        merge_debug_all_df = pd.DataFrame()

    return (
        result_df,
        fail_detail_map,
        enriched_dataset_map,
        merge_debug_all_df
    )


#將所有 failed rules 合併為統一報表（橫向格式）。
def build_unified_fail_report_df(result_df, fail_detail_map, raw_spec):
    """
    規則：
      - 一列 = 一筆 deduplicated fail record
      - item-level → 保留 ITEM
      - non-item-level → ITEM 為空
      - RCOL/RVALUE 橫向展開
    """

    if result_df is None or result_df.empty:
        return pd.DataFrame()

    rows = []

    # -------------------------------------------------
    # variable metadata
    # -------------------------------------------------
    var_df = raw_spec.get("variable") if isinstance(raw_spec, dict) else pd.DataFrame()

    if var_df is None or var_df.empty:
        var_df = pd.DataFrame()
    else:
        var_df = normalize_upper_cols(var_df)

    var_label_map = {}
    var_order_map = {}

    if not var_df.empty:
        for _, rr in var_df.iterrows():

            v = str(rr.get("VARIABLE", "")).strip().upper()
            label = str(rr.get("LABEL", "")).strip()
            order = rr.get("VARIABLE_ORDER")

            if v:
                var_label_map[v] = label if label else v

                try:
                    var_order_map[v] = int(order)
                except Exception:
                    var_order_map[v] = 999999

    # -------------------------------------------------
    # iterate rules
    # -------------------------------------------------
    for _, r in result_df.iterrows():

        if r.get("FAIL_COUNT", 0) <= 0:
            continue

        idx = r.get("IDX")
        detail = fail_detail_map.get(idx, {})

        fail_df = detail.get("FAIL_DF", pd.DataFrame())

        if fail_df is None or fail_df.empty:
            continue

        fail_df = normalize_upper_cols(fail_df)

        subj_col = get_subject_key(fail_df)

        stat = str(r.get("STAT_DATASET", "")).strip().upper()

        is_item_level = bool(r.get("IS_ITEM_LEVEL"))

        target_vars = r.get("TARGET_VARS", []) or []
        ref_vars = r.get("REF_VARS", []) or []

        target_vars = [
            str(x).strip().upper()
            for x in target_vars
            if str(x).strip()
        ]

        ref_vars = [
            str(x).strip().upper()
            for x in ref_vars
            if str(x).strip()
        ]

        # TARGET 一定先放 RCOL1
        report_vars = target_vars + ref_vars
        report_vars = list(dict.fromkeys(report_vars))

        # -------------------------------------------------
        # DTC：只從 target STAT_DATASET 的 variable list 找
        # -------------------------------------------------
        dtc_var = None

        target_var_meta = pd.DataFrame()

        if not var_df.empty and "STAT_DATASET" in var_df.columns:
            target_var_meta = var_df[
                var_df["STAT_DATASET"]
                .astype(str)
                .str.strip()
                .str.upper()
                .eq(stat)
            ].copy()
        else:
            target_var_meta = var_df.copy()

        # 1) 優先 target_vars 中的 date-like variable
        target_dtc_candidates = [
            v for v in target_vars
            if (
                _is_eval_date_like_column(v)
                and v in fail_df.columns
            )
        ]

        if target_dtc_candidates:
            dtc_var = sorted(
                target_dtc_candidates,
                key=lambda x: var_order_map.get(x, 999999)
            )[0]

        else:
            # 2) fallback：target STAT_DATASET variable list 中第一個 date-like
            if not target_var_meta.empty and "VARIABLE" in target_var_meta.columns:

                target_date_candidates = []

                for _, vr in target_var_meta.iterrows():

                    v = str(vr.get("VARIABLE", "")).strip().upper()

                    if (
                        v
                        and v in fail_df.columns
                        and _is_eval_date_like_column(v)
                    ):
                        try:
                            order = int(vr.get("VARIABLE_ORDER"))
                        except Exception:
                            order = var_order_map.get(v, 999999)

                        target_date_candidates.append((v, order))

                if target_date_candidates:
                    dtc_var = sorted(
                        target_date_candidates,
                        key=lambda x: x[1]
                    )[0][0]

        # -------------------------------------------------
        # build rows
        # -------------------------------------------------
        for _, fr in fail_df.iterrows():

            row = {}

            # -----------------------------------------
            # 基本欄位
            # -----------------------------------------
            row["PTNO"] = fr.get(subj_col, pd.NA) if subj_col else pd.NA
            row["RULE"] = r.get("RULE")

            crf_val = r.get("FORM", "")
            if pd.isna(crf_val) or str(crf_val).strip().upper() in ["NONE", "NAN", "<NA>"]:
                crf_val = ""
            row["CRF"] = crf_val

            row["Visit"] = fr.get("VISIT", pd.NA)

            note_val = r.get("MESSAGE", "")
            if pd.isna(note_val) or str(note_val).strip().upper() in ["NONE", "NAN", "<NA>"]:
                note_val = ""
            row["NOTE"] = note_val

            pd_val = r.get("PD", "")
            if pd.isna(pd_val) or str(pd_val).strip().upper() in ["NONE", "NAN", "<NA>"]:
                pd_val = ""
            row["PD"] = pd_val

            # -----------------------------------------
            # TP 先空
            # -----------------------------------------
            row["TP"] = ""

            # -----------------------------------------
            # DTC
            # -----------------------------------------
            if dtc_var and dtc_var in fr.index:
                row["DTC"] = fr.get(dtc_var)
            else:
                row["DTC"] = pd.NA

            # -----------------------------------------
            # ITEM（只在 item-level）
            # -----------------------------------------
            item_val = ""

            if is_item_level:

                # 優先使用 FAIL_GRAIN_KEYS 裡面的 item key
                grain_keys = r.get("FAIL_GRAIN_KEYS", []) or []

                item_key_candidates = [
                    str(c).strip().upper()
                    for c in grain_keys
                    if (
                        str(c).strip().upper() not in [
                            str(subj_col).strip().upper() if subj_col else "",
                            "VISIT",
                            "VISIT_DVP"
                        ]
                        and re.search(
                            ITEM_KEY_SUFFIX_PATTERN,
                            str(c).strip().upper()
                        )
                    )
                ]

                # fallback：如果 result_df 沒帶 FAIL_GRAIN_KEYS，就掃 fail_df
                if not item_key_candidates:
                    item_key_candidates = [
                        c for c in fail_df.columns
                        if re.search(ITEM_KEY_SUFFIX_PATTERN, c)
                    ]

                for c in item_key_candidates:
                    if c in fr.index:
                        v = fr.get(c)

                        if pd.notna(v) and str(v).strip():
                            item_val = v
                            break

            row["ITEM"] = item_val

            # -----------------------------------------
            # RCOL / RVALUE（TARGET → REF）
            # -----------------------------------------
            i = 1

            for v in report_vars:

                if v not in fr.index:
                    continue

                val = fr.get(v)

                if pd.isna(val):
                    continue

                label = var_label_map.get(v, v)

                row[f"RCOL{i}"] = label
                row[f"RVALUE{i}"] = val

                i += 1

            rows.append(row)

    report_df = pd.DataFrame(rows)

    if report_df.empty:
        return report_df

    # -------------------------------------------------
    # 排序
    # -------------------------------------------------
    sort_cols = ["PTNO", "RULE", "Visit", "ITEM"]

    sort_cols = [c for c in sort_cols if c in report_df.columns]

    report_df = (
        report_df
        .sort_values(sort_cols, na_position="last")
        .reset_index(drop=True)
    )

    return report_df
