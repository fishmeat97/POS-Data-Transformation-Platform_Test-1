import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# UI SECTION HEADER
# ============================================================

def section_header(
    step_number,
    title,
    subtitle="",
):

    st.markdown(
f"""
<div style="
    background: linear-gradient(
        90deg,
        rgba(80,80,80,0.10),
        rgba(80,80,80,0.03)
    );
    padding: 18px 22px;
    border-radius: 14px;
    border: 1px solid rgba(128,128,128,0.18);
    margin-top: 24px;
    margin-bottom: 16px;
">
<div style="
    display: flex;
    align-items: center;
    gap: 14px;
">

<div style="
    width: 44px;
    height: 44px;
    min-width: 44px;
    border-radius: 50%;
    background: rgba(120,120,120,0.18);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 21px;
    font-weight: 700;
">
{step_number}
</div>

<div>

<div style="
    font-size: 23px;
    font-weight: 700;
    margin-bottom: 2px;
">
{title}
</div>

<div style="
    font-size: 14px;
    color: #888;
">
{subtitle}
</div>

</div>

</div>
</div>
""",
        unsafe_allow_html=True,
    )






# ============================================================
# Page Config
# ============================================================
import streamlit as st

st.set_page_config(
    page_title="POS Monthly Data Mapping",
    page_icon="🥃",
    layout="wide",
)




# ============================================================
# HEADER
# ============================================================

header_col1, header_col2 = st.columns(
    [0.8, 6],
    vertical_alignment="center",
    gap="small",
)

with header_col1:
    st.image(
        "logo.png",
        width=160,
    )

with header_col2:
    st.markdown(
        """
        <h1 style="
            margin-bottom: 0;
            padding-bottom: 0;
            margin-left: -20px;
        ">
            POS Monthly Data Mapping
        </h1>

        <p style="
            color: grey;
            font-size: 16px;
            margin-top: 4px;
            margin-left: -10px;
        ">
            Monthly POS Data Integration,
            Mapping & Submission Dashboard
        </p>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ============================================================
# Final Output Columns
# ============================================================
OUTPUT_COLUMNS = [
    "地區",
    "業務員",
    "合約等級",
    "合約編號",
    "合約名稱",
    "店家編號",
    "店家名稱",
    "銷量(瓶)",
    "單價",
    "總價",
    "建議售價",
    "價格帶",
    "製造商",
    "品牌",
    "統一品項名稱",
    "品項",
    "容量",
    "品類",
    "年",
    "月",
]


# ============================================================
# Mapping Required Fields
# ============================================================
CUSTOMER_REQUIRED = [
    "REGION",
    "Sales ID",
    "Sales",
    "CONTRACT TYPE",
    "Contract JDE",
    "Contract NAME",
    "Outlet No",
    "Outlet NAME",
    "Rawdata Name",
]

SKU_REQUIRED = [
    "Mapping Name",
    "RSP",
    "Price Band",
    "Manufacture",
    "Brand",
    "SKU Name",
    "SIZE",
    "CATEGORY",
]


# ============================================================
# General Helpers
# ============================================================
def norm_text(value):
    if pd.isna(value):
        return ""

    text = unicodedata.normalize("NFKC", str(value))
    text = text.strip()
    text = re.sub(r"\s+", " ", text)

    return text


def norm_key(value):
    text = norm_text(value).upper()
    text = re.sub(
        r"[^\w\u4e00-\u9fff]+",
        "",
        text,
        flags=re.UNICODE,
    )
    return text


def norm_col(value):
    text = norm_text(value).lower()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def clean_file_label(filename):
    stem = Path(filename).stem
    stem = re.sub(r"\(\d+\)$", "", stem).strip()
    return stem


# ============================================================
# Numeric Cleaner
#
# Supports:
# 26瓶        -> 26
# 2,408瓶     -> 2408
# 10 箱       -> 10
# 15支        -> 15
# 1,250.5瓶   -> 1250.5
# ============================================================

def numeric(series):

    cleaned = (
        series
        .astype(str)

        # Remove commas
        .str.replace(
            ",",
            "",
            regex=False,
        )

        # Remove common Chinese quantity units
        .str.replace(
            r"(瓶|支|罐|箱|盒|包|桶|組|入|個|PCS|PC|BOTTLE|BOTTLES)",
            "",
            regex=True,
            case=False,
        )

        # Remove whitespace
        .str.replace(
            r"\s+",
            "",
            regex=True,
        )

        # Keep only number-related characters
        .str.extract(
            r"([-+]?\d*\.?\d+)",
            expand=False,
        )
    )

    return pd.to_numeric(
        cleaned,
        errors="coerce",
    )


def norm_id_key(value):

    if pd.isna(value):
        return ""

    text = str(value).strip()

    # 31020549.0 -> 31020549
    if re.fullmatch(
        r"\d+\.0+",
        text,
    ):
        text = text.split(".")[0]

    text = re.sub(
        r"[^\w]+",
        "",
        text,
    )

    return text.upper()

# ============================================================
# Filename Parser
#
# Example:
# 2233_31035675_202607_31020549_加州洋酒三重店.xlsx
#
# Sales ID
# Contract JDE
# YYYYMM
# Outlet No
# Rawdata Name / Branch
# ============================================================
def parse_filename(filename):
    stem = clean_file_label(filename)

    parts = [
        p.strip()
        for p in stem.split("_")
        if p.strip()
    ]

    sales_id = parts[0] if len(parts) > 0 else ""
    contract_jde = parts[1] if len(parts) > 1 else ""
    period = parts[2] if len(parts) > 2 else ""
    outlet_no = parts[3] if len(parts) > 3 else ""
    rawdata_name = "_".join(parts[4:]) if len(parts) > 4 else ""

    year = None
    month = None

    digits = re.sub(r"\D", "", period)

    # YYYYMM
    if len(digits) == 6:
        y = int(digits[:4])
        m = int(digits[4:])

        if 1900 <= y <= 2200 and 1 <= m <= 12:
            year = y
            month = m

    # ROC date, e.g. 11509
    elif len(digits) == 5:
        y = int(digits[:3]) + 1911
        m = int(digits[3:])

        if 1 <= m <= 12:
            year = y
            month = m

    ok = all([
        sales_id,
        contract_jde,
        outlet_no,
        year,
        month,
    ])

    return {
        "FILE_SALES_ID": sales_id,
        "FILE_CONTRACT_JDE": contract_jde,
        "FILE_PERIOD": period,
        "FILE_OUTLET_NO": outlet_no,
        "FILE_RAWDATA_NAME": rawdata_name,
        "YEAR": year,
        "MONTH": month,
        "FILENAME_STATUS": "OK" if ok else "CHECK",
    }


# ============================================================
# File Readers
# ============================================================
def read_csv_flexible(uploaded_file, header=0):
    raw = uploaded_file.getvalue()
    last_error = None

    for encoding in [
        "utf-8-sig",
        "utf-8",
        "cp950",
        "big5",
        "latin1",
    ]:
        try:
            return pd.read_csv(
                io.BytesIO(raw),
                encoding=encoding,
                header=header,
            )
        except Exception as e:
            last_error = e

    raise ValueError(
        f"CSV 無法讀取：{last_error}"
    )


def read_excel_bytes(uploaded_file, **kwargs):
    uploaded_file.seek(0)
    return pd.read_excel(
        uploaded_file,
        **kwargs,
    )


def first_matching_column(df, keywords):
    columns = list(df.columns)

    # Exact match first
    for column in columns:
        current = norm_col(column)

        for keyword in keywords:
            if norm_col(keyword) == current:
                return column

    # Contains match
    for column in columns:
        current = norm_col(column)

        for keyword in keywords:
            keyword_normalized = norm_col(keyword)

            if (
                keyword_normalized
                and keyword_normalized in current
            ):
                return column

    return None








# ============================================================
# Base Output Records
# ============================================================
def base_records(
    df,
    meta,
    template_type,
    source_file,
):
    output = pd.DataFrame(index=df.index)

    output["SOURCE_FILE"] = source_file
    output["SOURCE_ROW"] = range(
        1,
        len(df) + 1,
    )
    output["TEMPLATE_TYPE"] = template_type

    for key, value in meta.items():
        output[key] = value

    return output


# ============================================================
# TEMPLATE 1
# Standard POS SKU Table
# Example: 冠德系列
# ============================================================
def transform_template_1(uploaded_file):
    meta = parse_filename(
        uploaded_file.name
    )

    suffix = Path(
        uploaded_file.name
    ).suffix.lower()

    if suffix == ".csv":
        df = read_csv_flexible(
            uploaded_file,
            header=0,
        )
    else:
        df = read_excel_bytes(
            uploaded_file,
            header=0,
        )

    df = (
        df
        .dropna(how="all")
        .reset_index(drop=True)
    )

    sku_col = first_matching_column(
        df,
        [
            "品項名稱 SKU Name",
            "SKU Name",
            "品項名稱",
            "品名",
            "商品名稱",
            "品名規格",
        ],
    )

    qty_col = first_matching_column(
        df,
        [
            "銷量 Quantity",
            "Quantity",
            "銷量",
            "數量",
            "瓶數",
            "淨銷數量",
        ],
    )

    unit_price_col = first_matching_column(
        df,
        [
            "單價",
            "Unit Price",
            "Price",
        ],
    )

    total_price_col = first_matching_column(
        df,
        [
            "總價",
            "Total Price",
            "Amount",
            "金額",
        ],
    )

    if (
        sku_col is None
        or qty_col is None
    ):
        raise ValueError(
            "Template 1 找不到必要欄位。"
            f" SKU={sku_col}, Quantity={qty_col}"
        )

    output = base_records(
        df,
        meta,
        "T1_STANDARD_POS",
        uploaded_file.name,
    )

    output["RAW_CUSTOMER"] = meta[
        "FILE_RAWDATA_NAME"
    ]

    output["RAW_ROW_CUSTOMER"] = ""

    output["RAW_SKU"] = (
        df[sku_col]
        .map(norm_text)
    )

    output["QTY"] = numeric(
        df[qty_col]
    )

    output["UNIT_PRICE"] = pd.NA
    

    output["TOTAL_PRICE"] = pd.NA
    

    brand_col = first_matching_column(
        df,
        [
            "品牌名稱 Brand Name",
            "品牌名稱",
            "Brand Name",
        ],
    )

    size_col = first_matching_column(
        df,
        [
            "規格 Size (ml)",
            "Size (ml)",
            "規格",
            "容量",
        ],
    )

    category_col = first_matching_column(
        df,
        [
            "品類 Category",
            "Category",
            "品類",
        ],
    )

    output["RAW_BRAND"] = (
        df[brand_col]
        if brand_col
        else ""
    )

    output["RAW_SIZE"] = (
        df[size_col]
        if size_col
        else ""
    )

    output["RAW_CATEGORY"] = (
        df[category_col]
        if category_col
        else ""
    )

    return output


# ============================================================
# TEMPLATE 2
# Horizontal Monthly Actual Sales
# Example: 洋酒城系列
# ============================================================
# ============================================================
# ============================================================
# ============================================================
# TEMPLATE 2
# Horizontal Monthly Actual Sales
#
# Example:
# 洋酒城系列
#
# 商品名稱 -> RAW_SKU
# 商品代號 -> RAW_PRODUCT_CODE
# 實銷     -> QTY
#
# Customer:
# 由檔名取得
#
# Price:
# 不使用 Template 價格
# Report 統一使用 SKU Mapping RSP
# ============================================================

def transform_template_2(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Read Excel
    # ========================================================

    raw = read_excel_bytes(
        uploaded_file,
        sheet_name=0,
        header=None,
    )


    if len(raw) < 4:

        raise ValueError(
            "Template 2 資料列不足。"
        )


    # ========================================================
    # 3. Find Header Row
    # ========================================================

    header_row = None


    for i in range(
        min(
            20,
            len(raw),
        )
    ):

        values = [
            norm_text(value)
            for value
            in raw.iloc[i].tolist()
        ]


        if (
            "商品代號" in values
            and
            "商品名稱" in values
        ):

            header_row = i

            break


    if header_row is None:

        raise ValueError(
            "Template 2 找不到「商品代號 / 商品名稱」表頭。"
        )


    # ========================================================
    # 4. Header + Date Row
    # ========================================================

    date_row = max(
        0,
        header_row - 1,
    )


    headers = [
        norm_text(value)
        for value
        in raw.iloc[
            header_row
        ].tolist()
    ]


    dates = list(
        raw.iloc[
            date_row
        ].tolist()
    )


    # ========================================================
    # 5. Forward Fill Date
    # ========================================================

    ff_dates = []

    current_date = None


    for value in dates:

        if (
            pd.notna(value)
            and
            norm_text(value)
        ):

            current_date = value


        ff_dates.append(
            current_date
        )


    # ========================================================
    # 6. Find Product Columns
    # ========================================================

    product_code_idx = next(
        (
            i
            for i, header
            in enumerate(headers)
            if header == "商品代號"
        ),
        None,
    )


    product_name_idx = next(
        (
            i
            for i, header
            in enumerate(headers)
            if header == "商品名稱"
        ),
        None,
    )


    if product_name_idx is None:

        raise ValueError(
            "Template 2 找不到「商品名稱」欄位。"
        )


    # ========================================================
    # 7. Target Reporting Period
    # ========================================================

    target_year = meta[
        "YEAR"
    ]

    target_month = meta[
        "MONTH"
    ]


    # ========================================================
    # 8. Find Correct Actual Sales Column
    # ========================================================

    actual_sales_idx = None


    for i, header in enumerate(
        headers
    ):

        if "實銷" not in header:
            continue


        date_value = pd.to_datetime(
            ff_dates[i],
            errors="coerce",
        )


        if (
            pd.notna(date_value)
            and
            target_year is not None
            and
            target_month is not None
            and
            date_value.year == target_year
            and
            date_value.month == target_month
        ):

            actual_sales_idx = i

            break


    # ========================================================
    # 9. Fallback
    # ========================================================

    if actual_sales_idx is None:

        actual_sales_idx = next(
            (
                i
                for i, header
                in enumerate(headers)
                if "實銷" in header
            ),
            None,
        )


    if actual_sales_idx is None:

        raise ValueError(
            "Template 2 找不到「實銷」欄位。"
        )


    # ========================================================
    # 10. Extract Data
    # ========================================================

    data = (
        raw
        .iloc[
            header_row + 1:
        ]
        .copy()
        .dropna(
            how="all"
        )
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 11. Raw SKU
    # ========================================================

    raw_sku = (
        data.iloc[
            :,
            product_name_idx
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # 12. Quantity
    # ========================================================

    qty = numeric(
        data.iloc[
            :,
            actual_sales_idx
        ]
    )


    # ========================================================
    # 13. Keep Valid Product Rows
    # ========================================================

    valid = (
        raw_sku
        .ne("")
    )


    data = (
        data.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    raw_sku = (
        raw_sku.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    qty = (
        qty.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 14. Build Standard Output
    # ========================================================

    output = base_records(
        data,
        meta,
        "T2_MONTHLY_ACTUAL",
        uploaded_file.name,
    )


    # ========================================================
    # 15. Customer Metadata
    #
    # Customer Mapping 之後會使用：
    #
    # FILE_OUTLET_NO
    # FILE_RAWDATA_NAME
    #
    # 兩者都已經由 base_records(meta) 保留
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    # ========================================================
    # 16. SKU
    # ========================================================

    output[
        "RAW_SKU"
    ] = raw_sku


    # ========================================================
    # 17. Quantity
    # ========================================================

    output[
        "QTY"
    ] = qty


    # ========================================================
    # 18. Price
    #
    # 不使用 Template Price
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 19. Product Code
    # ========================================================

    if product_code_idx is not None:

        output[
            "RAW_PRODUCT_CODE"
        ] = (
            data.iloc[
                :,
                product_code_idx
            ]
            .map(
                norm_text
            )
            .values
        )

    else:

        output[
            "RAW_PRODUCT_CODE"
        ] = ""


    # ========================================================
    # 20. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = "實銷"


    output[
        "SOURCE_PRICE_COLUMN"
    ] = "RSP_FROM_SKU_MAPPING"


    return output

# ============================================================
# TEMPLATE 3A
# Multi-sheet Transaction Detail
# Example: 國泰
# ============================================================
def transform_template_3_multisheet(uploaded_file):
    meta = parse_filename(
        uploaded_file.name
    )

    uploaded_file.seek(0)

    excel_file = pd.ExcelFile(
        uploaded_file
    )

    frames = []

    for sheet in excel_file.sheet_names:
        uploaded_file.seek(0)

        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet,
            header=0,
        )

        df = (
            df
            .dropna(how="all")
            .reset_index(drop=True)
        )

        sku_col = first_matching_column(
            df,
            ["品名規格"],
        )

        qty_col = first_matching_column(
            df,
            [
                "數量2",
                "數量",
            ],
        )

        customer_col = first_matching_column(
            df,
            ["名稱"],
        )

        unit_price_col = first_matching_column(
            df,
            [
                "單價",
                "價格",
            ],
        )

        total_price_col = first_matching_column(
            df,
            [
                "總價",
                "金額",
            ],
        )

        if (
            sku_col is None
            or qty_col is None
        ):
            continue

        temp = pd.DataFrame()

        temp["RAW_SKU"] = (
            df[sku_col]
            .map(norm_text)
        )

        temp["QTY"] = numeric(
            df[qty_col]
        )

        temp["RAW_ROW_CUSTOMER"] = (
            df[customer_col].map(norm_text)
            if customer_col is not None
            else ""
        )

        temp["UNIT_PRICE"] = (
            numeric(df[unit_price_col])
            if unit_price_col
            else pd.NA
        )

        temp["TOTAL_PRICE"] = (
            numeric(df[total_price_col])
            if total_price_col
            else pd.NA
        )

        temp["SOURCE_SHEET"] = sheet

        temp["SOURCE_INNER_ROW"] = range(
            2,
            len(df) + 2,
        )

        frames.append(temp)

    if not frames:
        raise ValueError(
            "Template 3 多工作表找不到「品名規格 / 數量2」。"
        )

    data = pd.concat(
        frames,
        ignore_index=True,
    )

    data = (
        data[
            data["RAW_SKU"].ne("")
        ]
        .reset_index(drop=True)
    )

    output = base_records(
        data,
        meta,
        "T3_TRANSACTION_MULTISHEET",
        uploaded_file.name,
    )

    output["RAW_CUSTOMER"] = meta[
        "FILE_RAWDATA_NAME"
    ]

    output["RAW_ROW_CUSTOMER"] = data[
        "RAW_ROW_CUSTOMER"
    ]

    output["RAW_SKU"] = data[
        "RAW_SKU"
    ]

    output["QTY"] = data[
        "QTY"
    ]

    output["UNIT_PRICE"] = data[
        "UNIT_PRICE"
    ]

    output["TOTAL_PRICE"] = data[
        "TOTAL_PRICE"
    ]

    output["SOURCE_SHEET"] = data[
        "SOURCE_SHEET"
    ]

    output["SOURCE_INNER_ROW"] = data[
        "SOURCE_INNER_ROW"
    ]

    return output


# ============================================================
# TEMPLATE 3B
# Single-sheet Sales Detail
# Example: 加州洋酒
# ============================================================
def transform_template_3_salesdetail(uploaded_file):
    meta = parse_filename(
        uploaded_file.name
    )

    raw = read_excel_bytes(
        uploaded_file,
        sheet_name=0,
        header=None,
    )

    header_row = None

    for i in range(
        min(
            30,
            len(raw),
        )
    ):
        values = [
            norm_text(x)
            for x in raw.iloc[i].tolist()
        ]

        has_product = (
            "貨品名稱" in values
        )

        has_qty = any(
            "數" in x
            and "量" in x
            for x in values
        )

        if (
            has_product
            and has_qty
        ):
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            "Template 3 銷貨明細找不到表頭。"
        )

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=header_row,
    )

    df.columns = [
        norm_text(c)
        for c in df.columns
    ]

    df = (
        df
        .dropna(how="all")
        .reset_index(drop=True)
    )

    sku_col = first_matching_column(
        df,
        [
            "貨品名稱",
        ],
    )

    qty_col = first_matching_column(
        df,
        [
            "數 量",
            "數量",
        ],
    )

    unit_price_col = first_matching_column(
        df,
        [
            "單 價",
            "單價",
        ],
    )

    total_price_col = first_matching_column(
        df,
        [
            "總 價",
            "總價",
        ],
    )

    customer_col = first_matching_column(
        df,
        ["客戶名稱"],
    )

    if (
        sku_col is None
        or qty_col is None
    ):
        raise ValueError(
            "Template 3 銷貨明細找不到貨品名稱或數量。"
        )

    valid = (
        df[sku_col]
        .map(norm_text)
        .ne("")
    )

    df = (
        df
        .loc[valid]
        .reset_index(drop=True)
    )

    output = base_records(
        df,
        meta,
        "T3_SALES_DETAIL",
        uploaded_file.name,
    )

    output["SOURCE_ROW"] = range(
        header_row + 2,
        header_row + 2 + len(df),
    )

    output["RAW_CUSTOMER"] = meta[
        "FILE_RAWDATA_NAME"
    ]

    output["RAW_ROW_CUSTOMER"] = (
        df[customer_col].map(norm_text)
        if customer_col
        else ""
    )

    output["RAW_SKU"] = (
        df[sku_col]
        .map(norm_text)
    )

    output["QTY"] = numeric(
        df[qty_col]
    )

    output["UNIT_PRICE"] = (
        numeric(df[unit_price_col])
        if unit_price_col
        else pd.NA
    )

    output["TOTAL_PRICE"] = (
        numeric(df[total_price_col])
        if total_price_col
        else pd.NA
    )

    return output


# ============================================================
# TEMPLATE 4
# Multi-page Product Sales Report
# Example: 麥金酒
#
# 商品貨號 -> RAW_PRODUCT_CODE
# 商品名稱 -> RAW_SKU
# 商品名稱右側欄位 -> QTY
#
# Customer:
# 從檔名取得
#
# Price:
# 不使用 Template 原始價格
# Report 統一使用 SKU Mapping RSP
#
# Special:
# 商品名稱可能因換頁 / 換行而拆成下一列
# 系統會自動 append 回上一筆商品
# ============================================================

def transform_template_4(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Read All Sheets
    # ========================================================

    uploaded_file.seek(0)

    sheets = pd.read_excel(
        uploaded_file,
        sheet_name=None,
        header=None,
    )


    if not isinstance(
        sheets,
        dict,
    ):
        raise ValueError(
            "Template 4 無法讀取 Excel 工作表。"
        )


    # ========================================================
    # 3. Store All Product Records
    # ========================================================

    records = []

    last_record_index = None


    # ========================================================
    # 4. Process Every Sheet
    # ========================================================

    for sheet_name, raw in sheets.items():

        if (
            raw is None
            or raw.empty
        ):
            continue


        raw = raw.copy()


        # ====================================================
        # 5. Find Header Row
        #
        # Expected:
        # 商品貨號 | 商品名稱 | QTY
        # ====================================================

        header_row = None


        for i in range(
            min(
                20,
                len(raw),
            )
        ):

            values = [
                norm_text(value)
                for value
                in raw.iloc[i].tolist()
            ]


            if (
                "商品貨號" in values
                and
                "商品名稱" in values
            ):

                header_row = i
                break


        if header_row is None:
            continue


        # ====================================================
        # 6. Get Header Positions
        # ====================================================

        headers = [
            norm_text(value)
            for value
            in raw.iloc[
                header_row
            ].tolist()
        ]


        product_code_idx = next(
            (
                i
                for i, header
                in enumerate(headers)
                if header == "商品貨號"
            ),
            None,
        )


        product_name_idx = next(
            (
                i
                for i, header
                in enumerate(headers)
                if header == "商品名稱"
            ),
            None,
        )


        if (
            product_code_idx is None
            or product_name_idx is None
        ):
            continue


        # ====================================================
        # 7. QTY Column
        #
        # 麥金酒的數量欄沒有固定 Header
        # 數量在「商品名稱」右邊第一欄
        # ====================================================

        qty_idx = (
            product_name_idx
            + 1
        )


        if qty_idx >= raw.shape[1]:

            raise ValueError(
                f"Template 4｜{sheet_name} 找不到數量欄位。"
            )


        # ====================================================
        # 8. Process Data Rows
        # ====================================================

        for row_idx in range(
            header_row + 1,
            len(raw),
        ):

            row = raw.iloc[
                row_idx
            ]


            # ------------------------------------------------
            # Product Code
            # ------------------------------------------------

            raw_product_code = ""

            if (
                product_code_idx
                < len(row)
            ):

                raw_product_code = norm_text(
                    row.iloc[
                        product_code_idx
                    ]
                )


            # ------------------------------------------------
            # SKU Name
            # ------------------------------------------------

            raw_sku = ""

            if (
                product_name_idx
                < len(row)
            ):

                raw_sku = norm_text(
                    row.iloc[
                        product_name_idx
                    ]
                )


            # ------------------------------------------------
            # Quantity
            # ------------------------------------------------

            qty_value = pd.NA

            if qty_idx < len(row):

                qty_value = pd.to_numeric(
                    row.iloc[
                        qty_idx
                    ],
                    errors="coerce",
                )


            # =================================================
            # CASE 1
            # Normal Product Row
            #
            # Example:
            # T001 | 黑牌12年700cc | 5
            # =================================================

            if (
                raw_sku != ""
                and
                pd.notna(
                    qty_value
                )
            ):

                records.append(
                    {
                        "SOURCE_FILE":
                            uploaded_file.name,

                        "SOURCE_SHEET":
                            sheet_name,

                        "SOURCE_ROW":
                            row_idx + 1,

                        "SOURCE_INNER_ROW":
                            row_idx + 1,

                        "RAW_PRODUCT_CODE":
                            raw_product_code,

                        "RAW_SKU":
                            raw_sku,

                        "QTY":
                            qty_value,
                    }
                )


                last_record_index = (
                    len(records)
                    - 1
                )


                continue


            # =================================================
            # CASE 2
            # Wrapped SKU / Continuation Row
            #
            # Example:
            #
            # T020-1 | 蘇格登13年貴腐酒桶(亞 | 23
            # 版)
            #
            # 下一列的「版)」會出現在第一欄
            # =================================================

            is_continuation = (

                raw_product_code != ""

                and

                raw_sku == ""

                and

                pd.isna(
                    qty_value
                )
            )


            if (
                is_continuation
                and
                last_record_index
                is not None
            ):

                continuation_text = (
                    raw_product_code
                )


                # 避免 footer / 無效文字
                if continuation_text not in [
                    "以下空白",
                    "合計",
                    "總計",
                ]:

                    previous_sku = (
                        records[
                            last_record_index
                        ][
                            "RAW_SKU"
                        ]
                    )


                    records[
                        last_record_index
                    ][
                        "RAW_SKU"
                    ] = (
                        previous_sku
                        +
                        continuation_text
                    )


    # ========================================================
    # 9. Validation
    # ========================================================

    if not records:

        raise ValueError(
            "Template 4 找不到有效商品資料。"
        )


    # ========================================================
    # 10. Build DataFrame
    # ========================================================

    data = pd.DataFrame(
        records
    )


    # ========================================================
    # 11. Remove Invalid Rows
    # ========================================================

    data = data[
        data[
            "RAW_SKU"
        ]
        .map(
            norm_text
        )
        .ne("")
    ].copy()


    data = (
        data
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 12. Build Standard Output
    # ========================================================

    output = base_records(
        data,
        meta,
        "T4_MAIJIN_MULTI_PAGE",
        uploaded_file.name,
    )


    # ========================================================
    # 13. Audit Fields
    # ========================================================

    output[
        "SOURCE_SHEET"
    ] = data[
        "SOURCE_SHEET"
    ].values


    output[
        "SOURCE_INNER_ROW"
    ] = data[
        "SOURCE_INNER_ROW"
    ].values


    output[
        "SOURCE_ROW"
    ] = data[
        "SOURCE_ROW"
    ].values


    # ========================================================
    # 14. Customer
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    # ========================================================
    # 15. Product
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = data[
        "RAW_PRODUCT_CODE"
    ].values


    output[
        "RAW_SKU"
    ] = data[
        "RAW_SKU"
    ].values


    # ========================================================
    # 16. Quantity
    # ========================================================

    output[
        "QTY"
    ] = pd.to_numeric(
        data[
            "QTY"
        ],
        errors="coerce",
    ).values


    # ========================================================
    # 17. Price
    #
    # 全部後續使用 SKU Mapping RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 18. Optional Raw Product Information
    # ========================================================

    output[
        "RAW_BRAND"
    ] = ""


    output[
        "RAW_SIZE"
    ] = ""


    output[
        "RAW_CATEGORY"
    ] = ""


    # ========================================================
    # 19. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = (
        "商品名稱右側欄位"
    )


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    return output



# ============================================================
# TEMPLATE 5
# POS Standard Table with Header on Row 2
# Example: 康尼克
#
# Header row:
# Excel Row 2
#
# SKU:
# 品項名稱 SKU Name
#
# QTY:
# 銷量 Quantity
#
# Brand:
# 品牌名稱 Brand Name
#
# Size:
# 規格 Size (ml)
#
# Category:
# 品類 Category
#
# Customer:
# From filename
#
# Price:
# Do not use source price
# Final report uses SKU Mapping RSP
# ============================================================

def transform_template_5(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Read Raw Excel
    # ========================================================

    uploaded_file.seek(0)

    raw = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=None,
    )


    if raw.empty:

        raise ValueError(
            "Template 5 檔案沒有資料。"
        )


    # ========================================================
    # 3. Find Header Row
    #
    # 不 hardcode row 2
    # 改成掃描前 10 列，比較安全
    # ========================================================

    header_row = None


    for i in range(
        min(
            10,
            len(raw),
        )
    ):

        values = [
            norm_text(value)
            for value
            in raw.iloc[i].tolist()
        ]


        has_sku = any(
            (
                "品項名稱" in value
                and
                "SKU Name" in value
            )
            or
            value == "品項名稱"
            for value
            in values
        )


        has_qty = any(
            (
                "銷量" in value
                and
                "Quantity" in value
            )
            or
            value == "銷量"
            for value
            in values
        )


        if (
            has_sku
            and
            has_qty
        ):

            header_row = i
            break


    if header_row is None:

        raise ValueError(
            "Template 5 找不到 "
            "「品項名稱 SKU Name / 銷量 Quantity」表頭。"
        )


    # ========================================================
    # 4. Read Again Using Correct Header
    # ========================================================

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=header_row,
    )


    # Clean headers
    df.columns = [
        norm_text(column)
        for column
        in df.columns
    ]


    df = (
        df
        .dropna(
            how="all"
        )
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 5. Find Required Columns
    # ========================================================

    sku_col = first_matching_column(
        df,
        [
            "品項名稱 SKU Name",
            "品項名稱",
            "SKU Name",
        ],
    )


    qty_col = first_matching_column(
        df,
        [
            "銷量 Quantity",
            "銷量",
            "Quantity",
        ],
    )


    if (
        sku_col is None
        or
        qty_col is None
    ):

        raise ValueError(
            "Template 5 找不到必要欄位。"
            f" SKU={sku_col}, Quantity={qty_col}"
        )


    # ========================================================
    # 6. Optional Columns
    # ========================================================

    brand_col = first_matching_column(
        df,
        [
            "品牌名稱 Brand Name",
            "品牌名稱",
            "Brand Name",
        ],
    )


    size_col = first_matching_column(
        df,
        [
            "規格 Size (ml)",
            "Size (ml)",
            "規格",
            "容量",
        ],
    )


    category_col = first_matching_column(
        df,
        [
            "品類 Category",
            "Category",
            "品類",
        ],
    )


    product_code_col = first_matching_column(
        df,
        [
            "品項代碼（保樂力加產品） SKU code (for PRT)",
            "SKU code (for PRT)",
            "品項代碼",
            "SKU code",
        ],
    )


    barcode_col = first_matching_column(
        df,
        [
            "產品Barcode",
            "Barcode",
        ],
    )


    # ========================================================
    # 7. Keep Valid SKU Rows
    # ========================================================

    valid = (
        df[
            sku_col
        ]
        .map(
            norm_text
        )
        .ne("")
    )


    df = (
        df.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 8. Build Standard Output
    # ========================================================

    output = base_records(
        df,
        meta,
        "T5_POS_HEADER_ROW_2",
        uploaded_file.name,
    )


    # ========================================================
    # 9. Correct Excel Source Row
    #
    # header_row is zero-based.
    #
    # Example:
    # header_row = 1
    # first data row = Excel row 3
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = range(
        header_row + 2,
        header_row + 2 + len(df),
    )


    output[
        "SOURCE_SHEET"
    ] = "POS公版"


    # ========================================================
    # 10. Customer
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    # ========================================================
    # 11. SKU
    # ========================================================

    output[
        "RAW_SKU"
    ] = (
        df[
            sku_col
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # 12. Quantity
    # ========================================================

    output[
        "QTY"
    ] = numeric(
        df[
            qty_col
        ]
    )


    # ========================================================
    # 13. Product Code
    # ========================================================

    if product_code_col is not None:

        output[
            "RAW_PRODUCT_CODE"
        ] = (
            df[
                product_code_col
            ]
            .map(
                norm_text
            )
        )

    else:

        output[
            "RAW_PRODUCT_CODE"
        ] = ""


    # ========================================================
    # 14. Barcode
    # ========================================================

    if barcode_col is not None:

        output[
            "RAW_BARCODE"
        ] = (
            df[
                barcode_col
            ]
            .map(
                norm_text
            )
        )

    else:

        output[
            "RAW_BARCODE"
        ] = ""


    # ========================================================
    # 15. Brand
    # ========================================================

    output[
        "RAW_BRAND"
    ] = (
        df[
            brand_col
        ].map(
            norm_text
        )
        if brand_col is not None
        else ""
    )


    # ========================================================
    # 16. Size
    # ========================================================

    output[
        "RAW_SIZE"
    ] = (
        df[
            size_col
        ].map(
            norm_text
        )
        if size_col is not None
        else ""
    )


    # ========================================================
    # 17. Category
    # ========================================================

    output[
        "RAW_CATEGORY"
    ] = (
        df[
            category_col
        ].map(
            norm_text
        )
        if category_col is not None
        else ""
    )


    # ========================================================
    # 18. Price
    #
    # Final price always comes from SKU Mapping RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 19. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = qty_col


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    return output




# ============================================================
# TEMPLATE 6
# No-header Transaction Detail
# Example: 松賀
#
# Fixed column positions:
#
# 0 日期
# 1 客戶號碼
# 2 客戶名稱
# 3 商品號碼
# 4 商品名稱
# 5 單價
# 6 數量
# 7 小計
#
# Customer:
# Raw row customer is available
# Final customer mapping still uses filename metadata first
#
# Price:
# Source price retained for audit only
# Final report still uses SKU Mapping RSP
# ============================================================

def transform_template_6(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Read XLS without header
    # ========================================================

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=None,
        engine="xlrd",
    )


    df = (
        df
        .dropna(how="all")
        .reset_index(drop=True)
    )


    if df.empty:

        raise ValueError(
            "Template 6 檔案沒有資料。"
        )


    # ========================================================
    # 3. Validate Column Count
    # ========================================================

    if df.shape[1] < 8:

        raise ValueError(
            "Template 6 欄位數不足。"
            f" Expected >= 8, Actual={df.shape[1]}"
        )


    # ========================================================
    # 4. Fixed Column Positions
    # ========================================================

    DATE_COL = 0
    CUSTOMER_NO_COL = 1
    CUSTOMER_NAME_COL = 2
    PRODUCT_CODE_COL = 3
    SKU_COL = 4
    UNIT_PRICE_COL = 5
    QTY_COL = 6
    TOTAL_PRICE_COL = 7


    # ========================================================
    # 5. Raw Fields
    # ========================================================

    raw_sku = (
        df.iloc[
            :,
            SKU_COL
        ]
        .map(
            norm_text
        )
    )


    qty = numeric(
        df.iloc[
            :,
            QTY_COL
        ]
    )


    # ========================================================
    # 6. Keep Valid Product Rows
    #
    # 商品名稱不能是空白
    # 數量允許後面再檢查
    # ========================================================

    valid = (
        raw_sku
        .ne("")
    )


    df = (
        df.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    raw_sku = (
        raw_sku.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    qty = (
        qty.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 7. Build Standard Output
    # ========================================================

    output = base_records(
        df,
        meta,
        "T6_NO_HEADER_TRANSACTION",
        uploaded_file.name,
    )


    # ========================================================
    # 8. Source Row
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = range(
        1,
        len(df) + 1,
    )


    output[
        "SOURCE_SHEET"
    ] = "Sheet1"


    # ========================================================
    # 9. Customer
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = (
        df.iloc[
            :,
            CUSTOMER_NAME_COL
        ]
        .map(
            norm_text
        )
    )


    output[
        "RAW_CUSTOMER_NO"
    ] = (
        df.iloc[
            :,
            CUSTOMER_NO_COL
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # 10. Product
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = (
        df.iloc[
            :,
            PRODUCT_CODE_COL
        ]
        .map(
            norm_text
        )
    )


    output[
        "RAW_SKU"
    ] = raw_sku


    # ========================================================
    # 11. Quantity
    # ========================================================

    output[
        "QTY"
    ] = qty


    # ========================================================
    # 12. Transaction Date
    # ========================================================

    output[
        "RAW_TRANSACTION_DATE"
    ] = (
        df.iloc[
            :,
            DATE_COL
        ]
    )


    # ========================================================
    # 13. Source Price
    #
    # Audit only
    # ========================================================

    output[
        "RAW_UNIT_PRICE"
    ] = numeric(
        df.iloc[
            :,
            UNIT_PRICE_COL
        ]
    )


    output[
        "RAW_TOTAL_PRICE"
    ] = numeric(
        df.iloc[
            :,
            TOTAL_PRICE_COL
        ]
    )


    # ========================================================
    # 14. Final Price
    #
    # Final Report still uses SKU Mapping RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 15. Optional Raw Mapping Fields
    # ========================================================

    output[
        "RAW_BRAND"
    ] = ""


    output[
        "RAW_SIZE"
    ] = ""


    output[
        "RAW_CATEGORY"
    ] = ""


    # ========================================================
    # 16. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = "Column 6｜數量"


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    return output



# ============================================================
# Find Header Row by Standard Header Mapping
# ============================================================

def find_header_row(
    raw,
    required_fields,
    max_rows=30,
):

    for i in range(
        min(
            max_rows,
            len(raw),
        )
    ):

        values = [
            norm_text(value)
            for value
            in raw.iloc[i].tolist()
        ]

        all_found = True

        for field_name in required_fields:

            candidates = (
                STANDARD_HEADER_MAPPING.get(
                    field_name,
                    [],
                )
            )

            field_found = any(
                any(
                    norm_key(candidate)
                    ==
                    norm_key(value)
                    for candidate
                    in candidates
                )
                for value
                in values
                if norm_text(value)
            )

            if not field_found:
                all_found = False
                break

        if all_found:
            return i

    return None


# ============================================================
# TEMPLATE 7
# Generic SKU + QTY Transaction Detail
# ============================================================

def transform_template_7(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Read Raw Excel
    # ========================================================

    uploaded_file.seek(0)

    raw = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=None,
    )


    if raw.empty:

        raise ValueError(
            "Template 7 檔案沒有資料。"
        )


    # ========================================================
    # 3. Find Header Row
    # ========================================================

    header_row = find_header_row(
        raw,
        required_fields=[
            "SKU",
            "QTY",
        ],
        max_rows=30,
    )


    if header_row is None:

        raise ValueError(
            "Template 7 找不到 SKU / QTY 類型表頭。"
        )


    # ========================================================
    # 4. Read Again Using Correct Header
    # ========================================================

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=header_row,
    )


    df.columns = [
        norm_text(column)
        for column
        in df.columns
    ]


    df = (
        df
        .dropna(how="all")
        .reset_index(drop=True)
    )


    # ========================================================
    # 5. Find Standard Columns
    # ========================================================

    sku_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["SKU"],
    )


    qty_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["QTY"],
    )


    product_code_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["PRODUCT_CODE"],
    )


    customer_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["CUSTOMER"],
    )


    unit_price_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["UNIT_PRICE"],
    )


    total_price_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["TOTAL_PRICE"],
    )


    brand_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["BRAND"],
    )


    size_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["SIZE"],
    )


    category_col = first_matching_column(
        df,
        STANDARD_HEADER_MAPPING["CATEGORY"],
    )


    # ========================================================
    # 6. Validate Required Columns
    # ========================================================

    if (
        sku_col is None
        or
        qty_col is None
    ):

        raise ValueError(
            "Template 7 找不到必要欄位。"
            f" SKU={sku_col}, Quantity={qty_col}"
        )


    # ========================================================
    # 7. Prepare SKU / QTY
    # ========================================================

    raw_sku = (
        df[
            sku_col
        ]
        .map(
            norm_text
        )
    )


    qty = numeric(
        df[
            qty_col
        ]
    )


   # ========================================================
# 8. Keep Valid Product Rows
# ========================================================

    valid = (
        raw_sku
        .ne("")
    )


# ========================================================
# 駱北 ONLY
# 只保留「單位 = 瓶」
# ========================================================

    filename_text = norm_text(
        Path(
            uploaded_file.name
        ).stem
    )
    
    
    if "駱北" in filename_text:
    
        unit_col = first_matching_column(
            df,
            [
                "單位",
                "Unit",
            ],
        )
    
    
        if unit_col is None:
    
            raise ValueError(
                "Template 7｜駱北 找不到「單位」欄位。"
            )
    
    
        valid = (
            valid
            &
            df[unit_col]
            .map(norm_text)
            .eq("瓶")
        )
    
    
    # ========================================================
    # Apply Filter
    # ========================================================
    
    df = (
        df.loc[
            valid
        ]
        .reset_index(drop=True)
    )
    
    
    raw_sku = (
        raw_sku.loc[
            valid
        ]
        .reset_index(drop=True)
    )
    
    
    qty = (
        qty.loc[
            valid
        ]
        .reset_index(drop=True)
    )


    # ========================================================
    # 9. Create Output
    # ========================================================

    output = base_records(
        df,
        meta,
        "T7_TRANSACTION_DETAIL",
        uploaded_file.name,
    )


    # ========================================================
    # 10. Source Information
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = range(
        header_row + 2,
        header_row + 2 + len(df),
    )


    output[
        "SOURCE_SHEET"
    ] = "工作表1"


    # ========================================================
    # 11. Customer
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = (
        df[
            customer_col
        ]
        .map(
            norm_text
        )
        if customer_col is not None
        else ""
    )


    # ========================================================
    # 12. SKU
    # ========================================================

    output[
        "RAW_SKU"
    ] = raw_sku


    # ========================================================
    # 13. Quantity
    # ========================================================

    output[
        "QTY"
    ] = qty


    # ========================================================
    # 14. Product Code
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = (
        df[
            product_code_col
        ]
        .map(
            norm_text
        )
        if product_code_col is not None
        else ""
    )


    # ========================================================
    # 15. Brand
    # ========================================================

    output[
        "RAW_BRAND"
    ] = (
        df[
            brand_col
        ]
        .map(
            norm_text
        )
        if brand_col is not None
        else ""
    )


    # ========================================================
    # 16. Size
    # ========================================================

    output[
        "RAW_SIZE"
    ] = (
        df[
            size_col
        ]
        .map(
            norm_text
        )
        if size_col is not None
        else ""
    )


    # ========================================================
    # 17. Category
    # ========================================================

    output[
        "RAW_CATEGORY"
    ] = (
        df[
            category_col
        ]
        .map(
            norm_text
        )
        if category_col is not None
        else ""
    )


    # ========================================================
    # 18. Original Price
    # Audit only
    # ========================================================

    output[
        "RAW_UNIT_PRICE"
    ] = (
        numeric(
            df[
                unit_price_col
            ]
        )
        if unit_price_col is not None
        else pd.NA
    )


    output[
        "RAW_TOTAL_PRICE"
    ] = (
        numeric(
            df[
                total_price_col
            ]
        )
        if total_price_col is not None
        else pd.NA
    )


    # ========================================================
    # 19. Final Price
    # Final Report uses SKU Mapping RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 20. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = qty_col


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    return output

# ============================================================
# ============================================================
# ============================================================
# TEMPLATE 8
# 唯君系列
#
# Generic Multi-sheet Structure Parser
#
# Structure:
#
# Barcode
#   ↓
# Product Name / Specification area
#   ↓
# Unit
#   ↓
# Quantity
#   ↓
# Sales Amount
#
# IMPORTANT:
# SKU may span multiple columns.
# Therefore SKU is reconstructed from all valid text cells
# between BARCODE column and UNIT column.
# ============================================================

def transform_template_8(uploaded_file):

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 1. Read All Sheets
    # ========================================================

    uploaded_file.seek(0)

    sheets = pd.read_excel(
        uploaded_file,
        sheet_name=None,
        header=None,
    )


    if not sheets:

        raise ValueError(
            "Template 8 找不到任何工作表。"
        )


    frames = []


    # ========================================================
    # Helper: barcode validation
    # ========================================================

    def is_barcode(value):

        text = norm_text(
            value
        )

        if not text:
            return False


        # Excel sometimes reads barcode as 1234567890123.0
        text = re.sub(
            r"\.0$",
            "",
            text,
        )


        digits = re.sub(
            r"\D",
            "",
            text,
        )


        # 唯君商品條碼大多為 8~14 digits
        return (
            len(digits) >= 8
            and
            len(digits) <= 14
        )


    # ========================================================
    # Helper: clean barcode
    # ========================================================

    def clean_barcode(value):

        text = norm_text(
            value
        )

        text = re.sub(
            r"\.0$",
            "",
            text,
        )

        return text


    # ========================================================
    # Helper: identify numeric-only cell
    # ========================================================

    def is_pure_numeric(value):

        text = norm_text(
            value
        )

        if not text:
            return False


        cleaned = (
            text
            .replace(",", "")
            .replace("$", "")
        )


        number = pd.to_numeric(
            cleaned,
            errors="coerce",
        )


        return pd.notna(
            number
        )


    # ========================================================
    # Helper: build SKU from product area
    # ========================================================

    def build_sku_from_row(
        row,
        start_col,
        end_col,
    ):

        parts = []


        for col_idx in range(
            start_col,
            end_col,
        ):

            value = norm_text(
                row.iloc[
                    col_idx
                ]
            )


            if not value:
                continue


            # -----------------------------------------------
            # 排除純數字
            #
            # 例如：
            # 0.4
            # 700
            # 42.8
            #
            # 這些通常是 Excel 格式拆出的
            # 酒精濃度 / 規格數字
            #
            # 避免 RAW_SKU 變成奇怪字串
            # -----------------------------------------------

            if is_pure_numeric(
                value
            ):

                continue


            # -----------------------------------------------
            # 排除單位
            # -----------------------------------------------

            if value in [
                "瓶",
                "罐",
                "支",
                "箱",
                "盒",
            ]:

                continue


            parts.append(
                value
            )


        # -----------------------------------------------
        # Remove duplicate adjacent parts
        # -----------------------------------------------

        cleaned_parts = []

        for part in parts:

            if (
                not cleaned_parts
                or
                part
                !=
                cleaned_parts[-1]
            ):

                cleaned_parts.append(
                    part
                )


        return norm_text(
            " ".join(
                cleaned_parts
            )
        )


    # ========================================================
    # Process Every Sheet
    # ========================================================

    for (
        sheet_name,
        raw
    ) in sheets.items():

        if (
            raw is None
            or
            raw.empty
        ):

            continue


        # ====================================================
        # 2. Find Barcode Column
        #
        # Usually first column.
        # But don't assume.
        # ====================================================

        barcode_col = None


        best_barcode_score = 0


        for col_idx in range(
            raw.shape[1]
        ):

            series = raw.iloc[
                1:,
                col_idx
            ]


            if len(series) == 0:
                continue


            barcode_count = sum(
                is_barcode(value)
                for value
                in series
            )


            score = (
                barcode_count
                /
                max(
                    len(series),
                    1,
                )
            )


            if (
                score
                >
                best_barcode_score
            ):

                best_barcode_score = score

                barcode_col = col_idx


        if (
            barcode_col is None
            or
            best_barcode_score < 0.30
        ):

            continue


        # ====================================================
        # 3. Find Unit Column
        #
        # Find column containing mostly:
        # 瓶 / 罐 / 支 / 箱
        # ====================================================

        unit_col = None
        best_unit_score = 0


        unit_values = {
            "瓶",
    
        }


        for col_idx in range(
            barcode_col + 1,
            raw.shape[1]
        ):

            series = (
                raw.iloc[
                    1:,
                    col_idx
                ]
                .map(
                    norm_text
                )
            )


            nonempty = (
                series[
                    series.ne("")
                ]
            )


            if nonempty.empty:
                continue


            unit_count = (
                nonempty
                .isin(
                    unit_values
                )
                .sum()
            )


            score = (
                unit_count
                /
                len(
                    nonempty
                )
            )


            # Need enough actual rows
            if (
                score
                >
                best_unit_score
                and
                unit_count >= 3
            ):

                best_unit_score = score

                unit_col = col_idx


        if unit_col is None:

            continue


        # ====================================================
        # 4. Find QTY Column
        #
        # Search columns AFTER unit.
        #
        # QTY should:
        # - mostly numeric
        # - be close to UNIT
        # ====================================================

        qty_col = None


        for col_idx in range(
            unit_col + 1,
            raw.shape[1]
        ):

            test_numeric = numeric(
                raw.iloc[
                    1:,
                    col_idx
                ]
            )


            numeric_count = (
                test_numeric
                .notna()
                .sum()
            )


            total_rows = max(
                len(
                    test_numeric
                ),
                1,
            )


            numeric_ratio = (
                numeric_count
                /
                total_rows
            )


            # -----------------------------------------------
            # QTY should cover most product rows
            # -----------------------------------------------

            if numeric_ratio >= 0.50:

                qty_col = col_idx

                break


        if qty_col is None:

            continue


        # ====================================================
        # 5. Find Amount Column
        #
        # Usually final numeric-heavy column
        # ====================================================

        amount_col = None


        for col_idx in range(
            qty_col + 1,
            raw.shape[1]
        ):

            test_numeric = numeric(
                raw.iloc[
                    1:,
                    col_idx
                ]
            )


            numeric_ratio = (
                test_numeric
                .notna()
                .sum()
                /
                max(
                    len(
                        test_numeric
                    ),
                    1,
                )
            )


            if numeric_ratio >= 0.50:

                amount_col = col_idx


        # ====================================================
        # 6. Process Product Rows
        # ====================================================

        records = []


        for row_idx in range(
            1,
            len(raw),
        ):

            row = raw.iloc[
                row_idx
            ]


            barcode = row.iloc[
                barcode_col
            ]


            # -----------------------------------------------
            # Only real product rows
            # -----------------------------------------------

            if not is_barcode(
                barcode
            ):

                continue


            # -----------------------------------------------
            # Quantity
            # -----------------------------------------------

            qty_value = numeric(
                pd.Series([
                    row.iloc[
                        qty_col
                    ]
                ])
            ).iloc[0]


            if pd.isna(
                qty_value
            ):

                continue


            # =================================================
            # SKU
            #
            # Everything between barcode and unit,
            # but only valid text cells
            # =================================================

            sku = build_sku_from_row(
                row=row,
                start_col=barcode_col + 1,
                end_col=unit_col,
            )


            if not sku:

                continue


            # =================================================
            # Amount
            # =================================================

            if amount_col is not None:

                amount_value = numeric(
                    pd.Series([
                        row.iloc[
                            amount_col
                        ]
                    ])
                ).iloc[0]

            else:

                amount_value = pd.NA


            records.append({

                "RAW_PRODUCT_CODE":
                    clean_barcode(
                        barcode
                    ),

                "RAW_SKU":
                    sku,

                "QTY":
                    qty_value,

                "RAW_TOTAL_PRICE":
                    amount_value,

                "SOURCE_ROW":
                    row_idx + 1,

                "SOURCE_SHEET":
                    sheet_name,

                "SOURCE_BARCODE_COLUMN":
                    barcode_col,

                "SOURCE_UNIT_COLUMN":
                    unit_col,

                "SOURCE_QTY_COLUMN":
                    qty_col,

                "SOURCE_AMOUNT_COLUMN":
                    (
                        amount_col
                        if amount_col is not None
                        else ""
                    ),

            })


        # ====================================================
        # 7. Sheet Result
        # ====================================================

        if records:

            frames.append(
                pd.DataFrame(
                    records
                )
            )


    # ========================================================
    # 8. Validate
    # ========================================================

    if not frames:

        raise ValueError(
            "Template 8 所有 Sheet 都沒有找到有效的 "
            "商品條碼 / SKU / QTY 結構。"
        )


    # ========================================================
    # 9. Combine Sheets
    # ========================================================

    data = pd.concat(
        frames,
        ignore_index=True,
        sort=False,
    )


    # ========================================================
    # 10. Build Standard Output
    # ========================================================

    output = base_records(
        data,
        meta,
        "T8_WEIJUN_STRUCTURE",
        uploaded_file.name,
    )


    # ========================================================
    # Source
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = (
        data[
            "SOURCE_ROW"
        ]
    )


    output[
        "SOURCE_SHEET"
    ] = (
        data[
            "SOURCE_SHEET"
        ]
    )


    # ========================================================
    # Customer
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = (
        meta[
            "FILE_RAWDATA_NAME"
        ]
    )


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    # ========================================================
    # Product
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = (
        data[
            "RAW_PRODUCT_CODE"
        ]
    )


    output[
        "RAW_SKU"
    ] = (
        data[
            "RAW_SKU"
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # Quantity
    # ========================================================

    output[
        "QTY"
    ] = pd.to_numeric(
        data[
            "QTY"
        ],
        errors="coerce",
    )


    # ========================================================
    # Raw Amount
    #
    # 最終報表仍使用 RSP 計價
    # 這裡只保留來源銷貨金額供 audit
    # ========================================================

    output[
        "RAW_TOTAL_PRICE"
    ] = pd.to_numeric(
        data[
            "RAW_TOTAL_PRICE"
        ],
        errors="coerce",
    )


    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    output[
        "RAW_UNIT_PRICE"
    ] = pd.NA


    # ========================================================
    # Optional
    # ========================================================

    output[
        "RAW_BRAND"
    ] = ""


    output[
        "RAW_SIZE"
    ] = ""


    output[
        "RAW_CATEGORY"
    ] = ""


    # ========================================================
    # Audit
    # ========================================================

    output[
        "SOURCE_BARCODE_COLUMN"
    ] = (
        data[
            "SOURCE_BARCODE_COLUMN"
        ]
    )


    output[
        "SOURCE_UNIT_COLUMN"
    ] = (
        data[
            "SOURCE_UNIT_COLUMN"
        ]
    )


    output[
        "SOURCE_QTY_COLUMN"
    ] = (
        data[
            "SOURCE_QTY_COLUMN"
        ]
    )


    output[
        "SOURCE_AMOUNT_COLUMN"
    ] = (
        data[
            "SOURCE_AMOUNT_COLUMN"
        ]
    )


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    # ========================================================
    # IMPORTANT
    #
    # Do NOT put DataFrame inside output.attrs
    # ========================================================

    output.attrs = {}


    return output
       


# ============================================================
# TEMPLATE 9
# New Store Lee / Diageo Net Sales
# Example: 新店李系列
#
# Structure:
#
# Row 1:
# 帝亞吉歐
#
# Row 2:
# 品號 | 品名規格 | 淨銷數量
#
# Mapping:
# 品號     -> RAW_PRODUCT_CODE
# 品名規格 -> RAW_SKU
# 淨銷數量 -> QTY
#
# Customer:
# From filename
#
# Price:
# Final Report uses SKU Mapping RSP
# ============================================================

def transform_template_9(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Read Raw Excel
    # ========================================================

    uploaded_file.seek(0)

    raw = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=None,
    )


    if raw.empty:

        raise ValueError(
            "Template 9 檔案沒有資料。"
        )


    # ========================================================
    # 3. Find Header Row
    #
    # Expected:
    # 品號 | 品名規格 | 淨銷數量
    #
    # 不 hardcode row 2，
    # 掃描前 20 rows 比較安全
    # ========================================================

    header_row = None


    for i in range(
        min(
            20,
            len(raw),
        )
    ):

        values = [
            norm_text(value)
            for value
            in raw.iloc[i].tolist()
        ]


        has_product_code = any(
            norm_key(value)
            ==
            norm_key("品號")
            for value
            in values
        )


        has_sku = any(
            norm_key(value)
            ==
            norm_key("品名規格")
            for value
            in values
        )


        has_qty = any(
            norm_key(value)
            ==
            norm_key("淨銷數量")
            for value
            in values
        )


        if (
            has_product_code
            and
            has_sku
            and
            has_qty
        ):

            header_row = i
            break


    if header_row is None:

        raise ValueError(
            "Template 9 找不到"
            "「品號 / 品名規格 / 淨銷數量」表頭。"
        )


    # ========================================================
    # 4. Read Again Using Correct Header
    # ========================================================

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=header_row,
    )


    df.columns = [
        norm_text(column)
        for column
        in df.columns
    ]


    df = (
        df
        .dropna(
            how="all"
        )
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 5. Find Columns
    # ========================================================

    product_code_col = first_matching_column(
        df,
        [
            "品號",
        ],
    )


    sku_col = first_matching_column(
        df,
        [
            "品名規格",
        ],
    )


    qty_col = first_matching_column(
        df,
        [
            "淨銷數量",
        ],
    )


    # ========================================================
    # 6. Validate Required Columns
    # ========================================================

    if product_code_col is None:

        raise ValueError(
            "Template 9 找不到「品號」欄位。"
        )


    if sku_col is None:

        raise ValueError(
            "Template 9 找不到「品名規格」欄位。"
        )


    if qty_col is None:

        raise ValueError(
            "Template 9 找不到「淨銷數量」欄位。"
        )


    # ========================================================
    # 7. Prepare Raw SKU
    # ========================================================

    raw_sku = (
        df[
            sku_col
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # 8. Prepare Quantity
    #
    # numeric() should support:
    #
    # 42
    # 2408
    # 2,408
    # 26瓶
    # 2,408瓶
    # ========================================================

    qty = numeric(
        df[
            qty_col
        ]
    )


    # ========================================================
    # 9. Keep Valid Product Rows
    #
    # SKU 有值就保留
    # 不因 QTY 暫時轉換失敗而直接刪掉
    # ========================================================

    valid = (
        raw_sku
        .ne("")
    )


    df = (
        df.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    raw_sku = (
        raw_sku.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    qty = (
        qty.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 10. Build Standard Output
    # ========================================================

    output = base_records(
        df,
        meta,
        "T9_DIAGEO_NET_SALES",
        uploaded_file.name,
    )


    # ========================================================
    # 11. Source Information
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = range(
        header_row + 2,
        header_row + 2 + len(df),
    )


    output[
        "SOURCE_SHEET"
    ] = "工作表1"


    # ========================================================
    # 12. Customer
    #
    # From filename:
    #
    # 2172_31032115_202609_31032115_新店李系列.xlsx
    #
    # RAW_CUSTOMER = 新店李系列
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    # ========================================================
    # 13. Product Code
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = (
        df[
            product_code_col
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # 14. SKU
    # ========================================================

    output[
        "RAW_SKU"
    ] = raw_sku


    # ========================================================
    # 15. Quantity
    # ========================================================

    output[
        "QTY"
    ] = qty


    # ========================================================
    # 16. Source Quantity
    #
    # Audit 用：
    # 保留原本 Excel 的值
    # ========================================================

    output[
        "RAW_QTY"
    ] = (
        df[
            qty_col
        ]
        .astype(str)
    )


    # ========================================================
    # 17. Price
    #
    # Template 9 沒有價格欄
    #
    # Final Report:
    #
    # 單價 = RSP
    # 總價 = QTY * RSP
    # 建議售價 = RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    output[
        "RAW_UNIT_PRICE"
    ] = pd.NA


    output[
        "RAW_TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 18. Optional Raw Fields
    # ========================================================

    output[
        "RAW_BRAND"
    ] = ""


    output[
        "RAW_SIZE"
    ] = ""


    output[
        "RAW_CATEGORY"
    ] = ""


    # ========================================================
    # 19. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = qty_col


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    # ========================================================
    # 20. Return
    # ========================================================

    return output



# ============================================================
# TEMPLATE 10
# No Header - Fixed Two Columns
#
# Example:
# 名酒集系列
#
# Column 1 -> 商品 / SKU
# Column 2 -> 數量 / QTY
#
# NO HEADER
#
# Customer:
# From filename
#
# Price:
# Final Report uses SKU Mapping RSP
# ============================================================

def transform_template_10(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Read File
    # ========================================================

    uploaded_file.seek(0)

    suffix = Path(
        uploaded_file.name
    ).suffix.lower()


    if suffix == ".xls":

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            header=None,
            engine="xlrd",
        )

    else:

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            header=None,
        )


    # ========================================================
    # 3. Remove Empty Rows / Columns
    # ========================================================

    df = (
        df
        .dropna(
            how="all"
        )
        .dropna(
            axis=1,
            how="all",
        )
        .reset_index(
            drop=True
        )
    )


    if df.empty:

        raise ValueError(
            "Template 10 檔案沒有資料。"
        )


    # ========================================================
    # 4. Validate Column Count
    # ========================================================

    if df.shape[1] < 2:

        raise ValueError(
            "Template 10 欄位不足。"
            f" Expected >= 2, Actual={df.shape[1]}"
        )


    # ========================================================
    # 5. Fixed Column Positions
    #
    # Python index:
    #
    # Column 1 -> 0
    # Column 2 -> 1
    # ========================================================

    SKU_COL = 0
    QTY_COL = 1


    # ========================================================
    # 6. SKU
    # ========================================================

    raw_sku = (
        df.iloc[
            :,
            SKU_COL
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # 7. Quantity
    #
    # numeric() 可處理：
    #
    # 12
    # 1,200
    # 12瓶
    # 1,200瓶
    # ========================================================

    qty = numeric(
        df.iloc[
            :,
            QTY_COL
        ]
    )


    # ========================================================
    # 8. Keep Valid Rows
    #
    # 商品名稱有值才保留
    #
    # 同時排除很明顯不是商品資料的 row
    # ========================================================

    valid = (
        raw_sku
        .ne("")
    )


    df = (
        df.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    raw_sku = (
        raw_sku.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    qty = (
        qty.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 9. Build Standard Output
    # ========================================================

    output = base_records(
        df,
        meta,
        "T10_NO_HEADER_SKU_QTY",
        uploaded_file.name,
    )


    # ========================================================
    # 10. Source Info
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = range(
        1,
        len(df) + 1,
    )


    output[
        "SOURCE_SHEET"
    ] = "Sheet1"


    # ========================================================
    # 11. Customer
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    # ========================================================
    # 12. SKU
    # ========================================================

    output[
        "RAW_SKU"
    ] = raw_sku


    # ========================================================
    # 13. Quantity
    # ========================================================

    output[
        "QTY"
    ] = qty


    # ========================================================
    # 14. Keep Original Quantity for Audit
    # ========================================================

    output[
        "RAW_QTY"
    ] = (
        df.iloc[
            :,
            QTY_COL
        ]
        .astype(str)
    )


    # ========================================================
    # 15. Product Code
    #
    # Template 10 沒有 Product Code
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = ""


    # ========================================================
    # 16. Price
    #
    # Final Report:
    #
    # 單價 = SKU Mapping RSP
    # 總價 = QTY * RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    output[
        "RAW_UNIT_PRICE"
    ] = pd.NA


    output[
        "RAW_TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 17. Optional Fields
    # ========================================================

    output[
        "RAW_BRAND"
    ] = ""


    output[
        "RAW_SIZE"
    ] = ""


    output[
        "RAW_CATEGORY"
    ] = ""


    # ========================================================
    # 18. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = "Column 2"


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    # ========================================================
    # 19. Return
    # ========================================================

    return output


# ============================================================
# TEMPLATE 11
# 酒江村系列
#
# 國際條碼 -> RAW_PRODUCT_CODE
# 品牌     -> RAW_SKU
# 數量     -> QTY
# ============================================================

def transform_template_11(uploaded_file):

    meta = parse_filename(
        uploaded_file.name
    )

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=0,
    )

    df.columns = [
        norm_text(c)
        for c in df.columns
    ]

    product_code_col = first_matching_column(
        df,
        ["國際條碼"],
    )

    sku_col = first_matching_column(
        df,
        ["品牌"],
    )

    qty_col = first_matching_column(
        df,
        ["數量"],
    )

    if sku_col is None or qty_col is None:

        raise ValueError(
            "Template 11 找不到「品牌 / 數量」欄位。"
        )

    raw_sku = (
        df[sku_col]
        .map(norm_text)
    )

    qty = numeric(
        df[qty_col]
    )

    valid = (
        raw_sku.ne("")
    )

    df = (
        df.loc[valid]
        .reset_index(drop=True)
    )

    raw_sku = (
        raw_sku.loc[valid]
        .reset_index(drop=True)
    )

    qty = (
        qty.loc[valid]
        .reset_index(drop=True)
    )

    output = base_records(
        df,
        meta,
        "T11_JIUJIANGCUN",
        uploaded_file.name,
    )

    output["SOURCE_ROW"] = range(
        2,
        2 + len(df),
    )

    output["SOURCE_SHEET"] = "工作表1"

    output["RAW_CUSTOMER"] = (
        meta["FILE_RAWDATA_NAME"]
    )

    output["RAW_ROW_CUSTOMER"] = ""

    output["RAW_PRODUCT_CODE"] = (
        df[product_code_col]
        .map(norm_text)
        if product_code_col is not None
        else ""
    )

    output["RAW_SKU"] = raw_sku
    output["QTY"] = qty

    output["UNIT_PRICE"] = pd.NA
    output["TOTAL_PRICE"] = pd.NA

    output["RAW_BRAND"] = ""
    output["RAW_SIZE"] = ""
    output["RAW_CATEGORY"] = ""

    output["SOURCE_QTY_COLUMN"] = qty_col
    output["SOURCE_PRICE_COLUMN"] = (
        "RSP_FROM_SKU_MAPPING"
    )

    return output

# ============================================================
# TEMPLATE 12
# 廣茂隆 POS Public Format Variant
#
# SKU = 品項名稱 SKU Name
# QTY = 最後一欄 Column J
# ============================================================

def transform_template_12(uploaded_file):

    meta = parse_filename(
        uploaded_file.name
    )

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=0,
    )

    df.columns = [
        norm_text(c)
        for c in df.columns
    ]

    sku_col = first_matching_column(
        df,
        [
            "品項名稱 SKU Name",
            "品項名稱",
        ],
    )

    brand_col = first_matching_column(
        df,
        [
            "品牌名稱 Brand Name",
            "品牌名稱",
        ],
    )

    size_col = first_matching_column(
        df,
        [
            "規格 Size (ml)",
            "規格",
        ],
    )

    category_col = first_matching_column(
        df,
        [
            "品類 Category",
            "品類",
        ],
    )

    product_code_col = first_matching_column(
        df,
        [
            "品項代碼（保樂力加產品） SKU code (for PRT)",
            "SKU code",
        ],
    )

    if sku_col is None:

        raise ValueError(
            "Template 12 找不到品項名稱。"
        )

    # 最後一欄就是實際 QTY
    qty_col = df.columns[-1]

    raw_sku = (
        df[sku_col]
        .map(norm_text)
    )

    qty = numeric(
        df.iloc[:, -1]
    )

    # 這類公版很多 SKU 是清單但沒有銷量
    # 保留 SKU，但 QTY 沒值者後續仍可 audit
    valid = raw_sku.ne("")

    df = (
        df.loc[valid]
        .reset_index(drop=True)
    )

    raw_sku = (
        raw_sku.loc[valid]
        .reset_index(drop=True)
    )

    qty = (
        qty.loc[valid]
        .reset_index(drop=True)
    )

    output = base_records(
        df,
        meta,
        "T12_POS_BLANK_QTY_HEADER",
        uploaded_file.name,
    )

    output["SOURCE_ROW"] = range(
        2,
        2 + len(df),
    )

    output["SOURCE_SHEET"] = (
        str(pd.ExcelFile(uploaded_file).sheet_names[0])
        if False
        else "Sheet1"
    )

    output["RAW_CUSTOMER"] = (
        meta["FILE_RAWDATA_NAME"]
    )

    output["RAW_ROW_CUSTOMER"] = ""

    output["RAW_SKU"] = raw_sku
    output["QTY"] = qty

    output["RAW_PRODUCT_CODE"] = (
        df[product_code_col].map(norm_text)
        if product_code_col is not None
        else ""
    )

    output["RAW_BRAND"] = (
        df[brand_col].map(norm_text)
        if brand_col is not None
        else ""
    )

    output["RAW_SIZE"] = (
        df[size_col].map(norm_text)
        if size_col is not None
        else ""
    )

    output["RAW_CATEGORY"] = (
        df[category_col].map(norm_text)
        if category_col is not None
        else ""
    )

    output["UNIT_PRICE"] = pd.NA
    output["TOTAL_PRICE"] = pd.NA

    output["SOURCE_QTY_COLUMN"] = (
        "LAST_COLUMN"
    )

    output["SOURCE_PRICE_COLUMN"] = (
        "RSP_FROM_SKU_MAPPING"
    )

    return output
    
# ============================================================
# TEMPLATE 13
# 橡木桶 Monthly Horizontal Report
# ============================================================

def transform_template_13(uploaded_file):

    meta = parse_filename(
        uploaded_file.name
    )

    uploaded_file.seek(0)

    raw = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=None,
    )

    # ========================================================
    # Month from filename
    # ========================================================

    month_number = int(
        meta["MONTH"]
    )

    month_map = {
        1: "JAN",
        2: "FEB",
        3: "MAR",
        4: "APR",
        5: "MAY",
        6: "JUN",
        7: "JUL",
        8: "AUG",
        9: "SEP",
        10: "OCT",
        11: "NOV",
        12: "DEC",
    }

    target_month = month_map[
        month_number
    ]

    # ========================================================
    # Find month header row
    # ========================================================

    header_row = None
    qty_col_idx = None

    for i in range(
        min(20, len(raw))
    ):

        values = [
            norm_text(x).upper()
            for x
            in raw.iloc[i].tolist()
        ]

        if target_month in values:

            header_row = i

            qty_col_idx = (
                values.index(
                    target_month
                )
            )

            break

    if header_row is None:

        raise ValueError(
            f"Template 13 找不到月份欄位：{target_month}"
        )

    # ========================================================
    # Data:
    # Column A = Product Code
    # Column B = SKU
    # ========================================================

    product_code = (
        raw.iloc[
            header_row + 1:,
            0
        ]
        .map(norm_text)
        .reset_index(drop=True)
    )

    raw_sku = (
        raw.iloc[
            header_row + 1:,
            1
        ]
        .map(norm_text)
        .reset_index(drop=True)
    )

    qty = numeric(
        raw.iloc[
            header_row + 1:,
            qty_col_idx
        ]
        .reset_index(drop=True)
    )

    data = pd.DataFrame({
        "RAW_PRODUCT_CODE":
            product_code,

        "RAW_SKU":
            raw_sku,

        "QTY":
            qty,
    })

    # Remove TOTAL / empty rows
    valid = (
        data["RAW_SKU"].ne("")
        &
        ~data[
            "RAW_PRODUCT_CODE"
        ]
        .str.upper()
        .eq("TOTAL:")
    )

    data = (
        data.loc[valid]
        .reset_index(drop=True)
    )

    output = base_records(
        data,
        meta,
        "T13_MONTHLY_HORIZONTAL",
        uploaded_file.name,
    )

    output["SOURCE_SHEET"] = "Page1"

    output["SOURCE_ROW"] = range(
        header_row + 2,
        header_row + 2 + len(data),
    )

    output["RAW_CUSTOMER"] = (
        meta["FILE_RAWDATA_NAME"]
    )

    output["RAW_ROW_CUSTOMER"] = ""

    output["RAW_PRODUCT_CODE"] = (
        data["RAW_PRODUCT_CODE"]
    )

    output["RAW_SKU"] = (
        data["RAW_SKU"]
    )

    output["QTY"] = (
        data["QTY"]
    )

    output["UNIT_PRICE"] = pd.NA
    output["TOTAL_PRICE"] = pd.NA

    output["RAW_BRAND"] = ""
    output["RAW_SIZE"] = ""
    output["RAW_CATEGORY"] = ""

    output["SOURCE_QTY_COLUMN"] = (
        target_month
    )

    output["SOURCE_PRICE_COLUMN"] = (
        "RSP_FROM_SKU_MAPPING"
    )

    return output



# ============================================================
# TEMPLATE 14
# 美滿商行
#
# Fixed Column Structure:
#
# Column D -> SKU
# Column F -> Quantity
#
# Filter:
# - Remove blank SKU
# - Remove SKU cells starting with "D"
#
# Final Price:
# Use SKU Mapping RSP
# ============================================================

def transform_template_14(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Read XLS
    #
    # No header
    # ========================================================

    uploaded_file.seek(0)

    raw = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=None,
        engine="xlrd",
    )


    if raw.empty:

        raise ValueError(
            "Template 14 檔案沒有資料。"
        )


    # ========================================================
    # 3. Validate Column Count
    #
    # Excel:
    #
    # A = index 0
    # B = index 1
    # C = index 2
    # D = index 3 -> SKU
    # E = index 4
    # F = index 5 -> QTY
    # ========================================================

    if raw.shape[1] < 6:

        raise ValueError(
            "Template 14 欄位不足。"
            f"目前只有 {raw.shape[1]} 欄，"
            "至少需要到 F 欄。"
        )


    # ========================================================
    # 4. Fixed Columns
    # ========================================================

    SKU_COL = 3
    QTY_COL = 5


    # ========================================================
    # 5. Read SKU
    # ========================================================

    raw_sku = (
        raw.iloc[
            :,
            SKU_COL
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # 6. Read Quantity
    # ========================================================

    qty = numeric(
        raw.iloc[
            :,
            QTY_COL
        ]
    )


    # ========================================================
    # 7. Filter Valid SKU
    #
    # Keep:
    # - not blank
    #
    # Remove:
    # - blank
    # - starts with D
    #
    # Examples removed:
    #
    # D001-00-070-001
    # D002-12-070-003
    # Dxxxx...
    # ========================================================

    valid = (
        raw_sku.ne("")
        &
        ~raw_sku
        .str.strip()
        .str.upper()
        .str.startswith(
            "D"
        )
    )


    # ========================================================
    # 8. Apply Filter
    # ========================================================

    filtered = (
        raw.loc[
            valid
        ]
        .copy()
        .reset_index(
            drop=False
        )
    )


    raw_sku = (
        raw_sku.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    qty = (
        qty.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 9. Keep Original Excel Row Number
    #
    # index 0 -> Excel row 1
    # ========================================================

    source_rows = (
        filtered[
            "index"
        ]
        + 1
    )


    filtered = (
        filtered
        .drop(
            columns=[
                "index",
            ]
        )
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 10. Build Standard Output
    # ========================================================

    output = base_records(
        filtered,
        meta,
        "T14_MEIMAN_FIXED_D_F",
        uploaded_file.name,
    )


    # ========================================================
    # 11. Source Information
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = (
        source_rows
        .reset_index(
            drop=True
        )
    )


    output[
        "SOURCE_SHEET"
    ] = "Sheet1"


    # ========================================================
    # 12. Customer
    # ========================================================

    output[
        "RAW_CUSTOMER"
    ] = (
        meta[
            "FILE_RAWDATA_NAME"
        ]
    )


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    # ========================================================
    # 13. SKU
    #
    # Excel Column D
    # ========================================================

    output[
        "RAW_SKU"
    ] = raw_sku


    # ========================================================
    # 14. Quantity
    #
    # Excel Column F
    # ========================================================

    output[
        "QTY"
    ] = qty


    # ========================================================
    # 15. Original Quantity
    # ========================================================

    output[
        "RAW_QTY"
    ] = (
        filtered.iloc[
            :,
            QTY_COL
        ]
        .astype(str)
    )


    # ========================================================
    # 16. No Product Code
    #
    # D 開頭的商品代碼已經被排除
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = ""


    # ========================================================
    # 17. Price
    #
    # Final Report uses SKU Mapping RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    output[
        "RAW_UNIT_PRICE"
    ] = pd.NA


    output[
        "RAW_TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 18. Optional Raw Fields
    # ========================================================

    output[
        "RAW_BRAND"
    ] = ""


    output[
        "RAW_SIZE"
    ] = ""


    output[
        "RAW_CATEGORY"
    ] = ""


    # ========================================================
    # 19. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = "F"


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    # ========================================================
    # 20. Validate Result
    # ========================================================

    if output.empty:

        raise ValueError(
            "Template 14 過濾後沒有有效 SKU 資料。"
        )


    # ========================================================
    # 21. Return
    # ========================================================

    return output


# ============================================================
# TEMPLATE 15
# Multi-sheet SKU / QTY
# Example: 精湛酒藏
# ============================================================

def transform_template_15(uploaded_file):

    meta = parse_filename(
        uploaded_file.name
    )

    uploaded_file.seek(0)

    sheets = pd.read_excel(
        uploaded_file,
        sheet_name=None,
        header=0,
    )

    frames = []

    for (
        sheet_name,
        df
    ) in sheets.items():

        if df.empty:
            continue

        df.columns = [
            norm_text(c)
            for c
            in df.columns
        ]

        sku_col = first_matching_column(
            df,
            [
                "產品名稱",
                "商品名稱",
                "貨品名稱",
                "品項名稱",
                "品名",
            ],
        )

        qty_col = first_matching_column(
            df,
            [
                "銷售量",
                "數量",
                "銷量",
                "Quantity",
                "QTY",
            ],
        )

        if (
            sku_col is None
            or
            qty_col is None
        ):

            continue

        temp = pd.DataFrame()

        temp["RAW_SKU"] = (
            df[sku_col]
            .map(norm_text)
        )

        temp["QTY"] = numeric(
            df[qty_col]
        )

        temp["SOURCE_SHEET"] = (
            sheet_name
        )

        temp["SOURCE_ROW"] = range(
            2,
            2 + len(temp),
        )

        temp = temp[
            temp["RAW_SKU"]
            .ne("")
        ].copy()

        frames.append(
            temp
        )

    if not frames:

        raise ValueError(
            "Template 15 沒有找到可用的 SKU / QTY Sheet。"
        )

    data = pd.concat(
        frames,
        ignore_index=True,
    )

    output = base_records(
        data,
        meta,
        "T15_MULTI_SHEET_SKU_QTY",
        uploaded_file.name,
    )

    output["SOURCE_SHEET"] = (
        data["SOURCE_SHEET"]
    )

    output["SOURCE_ROW"] = (
        data["SOURCE_ROW"]
    )

    output["RAW_CUSTOMER"] = (
        meta["FILE_RAWDATA_NAME"]
    )

    output["RAW_ROW_CUSTOMER"] = ""

    output["RAW_PRODUCT_CODE"] = ""

    output["RAW_SKU"] = (
        data["RAW_SKU"]
    )

    output["QTY"] = (
        data["QTY"]
    )

    output["UNIT_PRICE"] = pd.NA
    output["TOTAL_PRICE"] = pd.NA

    output["RAW_BRAND"] = ""
    output["RAW_SIZE"] = ""
    output["RAW_CATEGORY"] = ""

    output["SOURCE_QTY_COLUMN"] = (
        "銷售量"
    )

    output["SOURCE_PRICE_COLUMN"] = (
        "RSP_FROM_SKU_MAPPING"
    )

    return output




# ============================================================
# Auto Detect Template Type
#
# Detection Priority:
#
# 1. Template 4 - 麥金酒 Multi-page / Multi-sheet
# 2. Template 5 - POS 公版 / 康尼克
# 3. Template 3A - 國泰 Multi-sheet
# 4. Template 2 - Horizontal Monthly Actual Sales
# 5. Template 3B - Single-sheet Sales Detail
# 6. Template 1 - Standard POS
# 7. UNKNOWN
# ============================================================

def detect_template_type(uploaded_file):

    suffix = Path(
        uploaded_file.name
    ).suffix.lower()

    filename_text = norm_text(
        Path(
            uploaded_file.name
        ).stem
    )


    # ========================================================
    # TEMPLATE 14
    # IMPORTANT:
    # 必須在任何 pd.ExcelFile / pd.read_excel 之前判斷
    #
    # 因為這是舊式 .xls 特殊報表
    # ========================================================

    if (
        suffix == ".xls"
        and
        "美滿商行"
        in filename_text
    ):

        return "TEMPLATE_14"


    # ========================================================
    # CSV
    # ========================================================

    if suffix == ".csv":

        try:

            uploaded_file.seek(0)

            df = read_csv_flexible(
                uploaded_file,
                header=0,
            )


            sku_col = first_matching_column(
                df,
                [
                    "品項名稱 SKU Name",
                    "SKU Name",
                    "品項名稱",
                    "品名",
                    "商品名稱",
                ],
            )


            qty_col = first_matching_column(
                df,
                [
                    "銷量 Quantity",
                    "Quantity",
                    "銷量",
                    "數量",
                    "瓶數",
                ],
            )


            if (
                sku_col is not None
                and
                qty_col is not None
            ):

                return "TEMPLATE_1"


        except Exception:
            pass


        return "UNKNOWN"


    # ========================================================
    # Excel Only
    # ========================================================

    if suffix not in [
        ".xlsx",
        ".xls",
    ]:

        return "UNKNOWN"


    # ========================================================
    # Load Excel Workbook
    # ========================================================

    try:

        uploaded_file.seek(0)

        excel_file = pd.ExcelFile(
            uploaded_file
        )

        sheet_names = (
            excel_file.sheet_names
        )

    except Exception:

        return "UNKNOWN"


    # ========================================================
    # Read First Sheet as RAW
    # ========================================================

    try:

        uploaded_file.seek(0)

        raw = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            header=None,
        )

    except Exception:

        return "UNKNOWN"


    # ========================================================
    # TEMPLATE 4
    # 麥金酒
    #
    # Characteristics:
    #
    # 商品貨號
    # 商品名稱
    #
    # Usually:
    # Multi-sheet / Page 1 / Page 2 / ...
    # ========================================================

    try:

        filename_text = norm_text(
            Path(
                uploaded_file.name
            ).stem
        )


        first_rows_text = " ".join(
            norm_text(value)
            for value
            in raw.iloc[
                :20
            ]
            .to_numpy()
            .flatten()
        )


        filename_is_maijin = (
            "麥金酒"
            in filename_text
        )


        has_maijin_title = (
            "麥金酒"
            in first_rows_text
        )


        has_sales_report_title = (
            "期間商品銷售統計表"
            in first_rows_text
        )


        has_template_4_header = False


        for i in range(
            min(
                30,
                len(raw),
            )
        ):

            values = [
                norm_text(value)
                for value
                in raw.iloc[
                    i
                ].tolist()
            ]


            if (
                "商品貨號" in values
                and
                "商品名稱" in values
            ):

                has_template_4_header = True
                break


        if (
            has_template_4_header
            and
            (
                filename_is_maijin
                or
                has_maijin_title
                or
                has_sales_report_title
            )
        ):

            return "TEMPLATE_4"


    except Exception:
        pass


    # ========================================================
    # TEMPLATE 5
    # POS 公版
    # Example: 康尼克
    #
    # Characteristics:
    #
    # Header 不一定在第 1 row
    #
    # 品牌名稱 Brand Name
    # 品項名稱 SKU Name
    # 規格 Size (ml)
    # 品類 Category
    # 銷量 Quantity
    # ========================================================

    try:

        has_template_5_header = False


        for i in range(
            min(
                15,
                len(raw),
            )
        ):

            values = [
                norm_text(value)
                for value
                in raw.iloc[
                    i
                ].tolist()
            ]


            has_sku = any(
                (
                    "品項名稱" in value
                    and
                    "SKU Name" in value
                )
                or
                value == "品項名稱"
                for value
                in values
            )


            has_qty = any(
                (
                    "銷量" in value
                    and
                    "Quantity" in value
                )
                or
                value == "銷量"
                for value
                in values
            )


            has_brand = any(
                (
                    "品牌名稱" in value
                    and
                    "Brand Name" in value
                )
                or
                value == "品牌名稱"
                for value
                in values
            )


            has_size = any(
                (
                    "規格" in value
                    and
                    "Size" in value
                )
                or
                value == "規格"
                or
                value == "容量"
                for value
                in values
            )


            has_category = any(
                (
                    "品類" in value
                    and
                    "Category" in value
                )
                or
                value == "品類"
                for value
                in values
            )


            if (
                has_sku
                and
                has_qty
                and
                has_brand
                and
                (
                    has_size
                    or
                    has_category
                )
            ):

                has_template_5_header = True
                break


        if has_template_5_header:

            return "TEMPLATE_5"


    except Exception:
        pass


# ========================================================
# TEMPLATE 6
# 松賀 No-header Transaction Detail
# ========================================================

    try:
    
        filename_text = norm_text(
            Path(
                uploaded_file.name
            ).stem
        )
    
    
        filename_is_songhe = (
            "松賀"
            in filename_text
        )
    
    
        has_expected_columns = (
            raw.shape[1] >= 8
        )
    
    
        if (
            filename_is_songhe
            and
            has_expected_columns
        ):
    
            return "TEMPLATE_6"
    
        
    except Exception:
        pass







    # ========================================================
# ========================================================
# TEMPLATE 8
# 唯君系列
# ========================================================

    filename_text = norm_text(
        Path(
            uploaded_file.name
        ).stem
    )
    
    
    if "唯君" in filename_text:
    
        return "TEMPLATE_8"

    
    # ========================================================
    # TEMPLATE 3A
    # 國泰 Multi-sheet
    #
    # Characteristics:
    #
    # Multiple sheets
    # 品名規格
    # 數量2 / 數量
    # ========================================================

    if len(
        sheet_names
    ) > 1:

        try:

            uploaded_file.seek(0)

            first_sheet = pd.read_excel(
                uploaded_file,
                sheet_name=sheet_names[
                    0
                ],
                header=0,
            )


            columns = {
                norm_col(column)
                for column
                in first_sheet.columns
            }


            has_product = any(
                "品名規格"
                in column
                for column
                in columns
            )


            has_qty = any(
                (
                    "數量2"
                    in column
                )
                or
                (
                    "數量"
                    in column
                )
                for column
                in columns
            )


            if (
                has_product
                and
                has_qty
            ):

                return "TEMPLATE_3"


        except Exception:
            pass


    # ========================================================
    # TEMPLATE 2
    # Horizontal Monthly Actual Sales
    #
    # Characteristics:
    #
    # 商品代號
    # 商品名稱
    # 實銷
    #
    # 商品 Header 與實銷不一定在完全同一列
    # ========================================================

    try:

        template_2_header_found = False

        actual_sales_found = False


        for i in range(
            min(
                30,
                len(raw),
            )
        ):

            values = [
                norm_text(value)
                for value
                in raw.iloc[
                    i
                ].tolist()
            ]


            if (
                "商品代號" in values
                and
                "商品名稱" in values
            ):

                template_2_header_found = True


            if any(
                "實銷"
                in value
                for value
                in values
            ):

                actual_sales_found = True


        if (
            template_2_header_found
            and
            actual_sales_found
        ):

            return "TEMPLATE_2"


    except Exception:
        pass


    # ========================================================
    # TEMPLATE 3B
    # Single-sheet Sales Detail
    #
    # Characteristics:
    #
    # 貨品名稱
    # 數量
    # ========================================================

    try:

        for i in range(
            min(
                40,
                len(raw),
            )
        ):

            values = [
                norm_text(value)
                for value
                in raw.iloc[
                    i
                ].tolist()
            ]


            has_product = (
                "貨品名稱"
                in values
            )


            has_qty = any(
                (
                    "數"
                    in value
                )
                and
                (
                    "量"
                    in value
                )
                for value
                in values
            )


            if (
                has_product
                and
                has_qty
            ):

                return "TEMPLATE_3"


    except Exception:
        pass


# ========================================================
# TEMPLATE 9
# 新店李系列 / Diageo Net Sales
#
# Expected structure:
#
# 品號 | 品名規格 | 淨銷數量
# ========================================================

    try:
    
        template_9_header_found = False
    
    
        for i in range(
            min(
                20,
                len(raw),
            )
        ):
    
            values = [
                norm_text(value)
                for value
                in raw.iloc[i].tolist()
            ]
    
    
            normalized_values = [
                norm_key(value)
                for value
                in values
                if norm_text(value)
            ]
    
    
            has_product_code = (
                norm_key("品號")
                in normalized_values
            )
    
    
            has_sku = (
                norm_key("品名規格")
                in normalized_values
            )
    
    
            has_qty = (
                norm_key("淨銷數量")
                in normalized_values
            )
    
    
            if (
                has_product_code
                and
                has_sku
                and
                has_qty
            ):
    
                template_9_header_found = True
                break
    
    
        if template_9_header_found:
    
            return "TEMPLATE_9"
    
    
    except Exception:
        pass



# ========================================================
# TEMPLATE 10
# 名酒集系列
# No Header / Fixed SKU + QTY Columns
# ========================================================

    try:
    
        filename_text = norm_text(
            Path(
                uploaded_file.name
            ).stem
        )
    
    
        filename_is_template_10 = (
            "名酒集"
            in filename_text
        )
    
    
        # 至少要有兩欄
        has_minimum_columns = (
            raw.shape[1] >= 2
        )
    
    
        if (
            filename_is_template_10
            and
            has_minimum_columns
        ):
    
            return "TEMPLATE_10"
    
    
    except Exception:
        pass

# ========================================================
# TEMPLATE 11
# 酒江村
# ========================================================

    try:
    
        normalized_columns = [
            norm_key(x)
            for x
            in raw.iloc[0].tolist()
            if norm_text(x)
        ]
    
        if (
            norm_key("國際條碼") in normalized_columns
            and
            norm_key("品牌") in normalized_columns
            and
            norm_key("數量") in normalized_columns
        ):
    
            return "TEMPLATE_11"
    
    except Exception:
        pass

# ========================================================
# TEMPLATE 12
# POS 公版 + Blank QTY Header
# ========================================================

    try:
    
        filename_text = norm_text(
            Path(uploaded_file.name).stem
        )
    
        first_row = [
            norm_key(x)
            for x in raw.iloc[0].tolist()
            if norm_text(x)
        ]
    
        if (
            "廣茂隆" in filename_text
            and
            any(
                norm_key("品項名稱")
                in x
                for x in first_row
            )
            and raw.shape[1] >= 10
        ):
    
            return "TEMPLATE_12"
    
    except Exception:
        pass


# ========================================================
# TEMPLATE 13
# 橡木桶
# ========================================================

    try:
    
        filename_text = norm_text(
            Path(uploaded_file.name).stem
        )
    
        first_20_text = " ".join(
            norm_text(x).upper()
            for x
            in raw.iloc[:20]
            .to_numpy()
            .flatten()
        )
    
        if (
            "橡木桶" in filename_text
            and
            "JAN" in first_20_text
            and
            "SEP" in first_20_text
            and
            "DEC" in first_20_text
        ):
    
            return "TEMPLATE_13"
    
    except Exception:
        pass
    
        


    # ========================================================
# TEMPLATE 15
# 精湛酒藏 Multi-sheet
# ========================================================

    try:
    
        filename_text = norm_text(
            Path(uploaded_file.name).stem
        )
    
        if (
            "精湛酒藏"
            in filename_text
            and
            len(sheet_names) > 1
        ):
    
            return "TEMPLATE_15"
    
    except Exception:
        pass


    # ========================================================
    # TEMPLATE 7
    # Generic SKU + QTY Template
    #
    # Examples:
    # 七七會社
    # 酒藏 / 艾姆度拉
    # 新店李系列
    #
    # Supported SKU headers:
    # 商品名稱
    # 貨品名稱
    # 品項名稱
    # 品名
    # 品名規格
    # 商品品名
    # 產品品名
    # 產品名稱
    #
    # Supported QTY headers:
    # 數量
    # 銷量
    # 淨銷數量
    # 實銷
    # Quantity
    # QTY
    # 瓶數
    # ========================================================

    try:
    
        template_7_header_row = find_header_row(
            raw,
            required_fields=[
                "SKU",
                "QTY",
            ],
            max_rows=30,
        )
    
    
        if template_7_header_row is not None:
    
            return "TEMPLATE_7"
    
    
    except Exception:
        pass
        


    
    # ========================================================
    # TEMPLATE 1
    # Standard POS
    #
    # Header must normally be Row 1
    # ========================================================

    try:

        uploaded_file.seek(0)

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            header=0,
        )


        sku_col = first_matching_column(
            df,
            [
                "品項名稱 SKU Name",
                "SKU Name",
                "品項名稱",
                "品名",
                "商品名稱",
            ],
        )


        qty_col = first_matching_column(
            df,
            [
                "銷量 Quantity",
                "Quantity",
                "銷量",
                "數量",
                "瓶數",
            ],
        )


        if (
            sku_col is not None
            and
            qty_col is not None
        ):

            return "TEMPLATE_1"


    except Exception:
        pass


    # ========================================================
    # UNKNOWN
    # ========================================================

    return "UNKNOWN"

# ============================================================
# Unified Auto Transformer
# ============================================================
# ============================================================
# Unified Auto Transformer
# ============================================================

# ============================================================
# Unified Auto Transformer
# ============================================================

def auto_transform_template(uploaded_file):

    uploaded_file.seek(0)

    template_type = detect_template_type(
        uploaded_file
    )

    uploaded_file.seek(0)


    # ========================================================
    # TEMPLATE 1
    # ========================================================

    if template_type == "TEMPLATE_1":

        result = transform_template_1(
            uploaded_file
        )

        return (
            result,
            "Template 1｜標準 PRT_POS Template",
        )


    # ========================================================
    # TEMPLATE 2
    # ========================================================

    elif template_type == "TEMPLATE_2":

        uploaded_file.seek(0)

        result = transform_template_2(
            uploaded_file
        )

        return (
            result,
            "Template 2｜月份實銷表",
        )


    # ========================================================
    # TEMPLATE 3
    # ========================================================

    elif template_type == "TEMPLATE_3":

        uploaded_file.seek(0)

        excel_file = pd.ExcelFile(
            uploaded_file
        )


        # Template 3A - Multi-sheet
        if len(
            excel_file.sheet_names
        ) > 1:

            try:

                uploaded_file.seek(0)

                result = (
                    transform_template_3_multisheet(
                        uploaded_file
                    )
                )

                return (
                    result,
                    "Template 3A｜多工作表銷售明細",
                )

            except Exception:
                pass


        # Template 3B - Single-sheet
        uploaded_file.seek(0)

        result = (
            transform_template_3_salesdetail(
                uploaded_file
            )
        )

        return (
            result,
            "Template 3B｜單工作表銷貨明細",
        )


    # ========================================================
    # TEMPLATE 4
    # 麥金酒
    # ========================================================

    elif template_type == "TEMPLATE_4":

        uploaded_file.seek(0)

        result = transform_template_4(
            uploaded_file
        )

        return (
            result,
            "Template 4｜麥金酒期間商品銷售統計表",
        )


    # ========================================================
    # TEMPLATE 5
    # POS 公版 / 康尼克
    # ========================================================

    elif template_type == "TEMPLATE_5":

        uploaded_file.seek(0)

        result = transform_template_5(
            uploaded_file
        )

        return (
            result,
            "Template 5｜POS 公版_HEADER on 2nd row",
        )





    elif template_type == "TEMPLATE_6":
    
        uploaded_file.seek(0)
    
        result = transform_template_6(
            uploaded_file
        )
    
        return (
            result,
            "Template 6｜(松賀) no HEADER",
        )
       



    
    elif template_type == "TEMPLATE_7":
    
        uploaded_file.seek(0)
    
        result = transform_template_7(
            uploaded_file
        )
    
        return (
            result,
            "Template 7｜(七七會社) HEADER on 2nd row",
        )
    


    elif template_type == "TEMPLATE_8":

        uploaded_file.seek(0)
    
        result = transform_template_8(
            uploaded_file
        )
        
        return (
            result,
            "Template 8｜唯君系列",
        )
            

    elif template_type == "TEMPLATE_9":
    
        uploaded_file.seek(0)
    
        result = transform_template_9(
            uploaded_file
        )
    
        return (
            result,
            "Template 9｜新店李淨銷報表",
        )



    elif template_type == "TEMPLATE_10":
    
        uploaded_file.seek(0)
    
        result = transform_template_10(
            uploaded_file
        )
    
        return (
            result,
            "Template 10｜名酒集無表頭 SKU / QTY",
        )


    elif template_type == "TEMPLATE_11":
    
        uploaded_file.seek(0)
    
        result = transform_template_11(
            uploaded_file
        )
    
        return (
            result,
            "Template 11｜酒江村商品銷量",
        )
    
    
    elif template_type == "TEMPLATE_12":
    
        uploaded_file.seek(0)
    
        result = transform_template_12(
            uploaded_file
        )
    
        return (
            result,
            "Template 12｜POS 公版無 QTY Header",
        )
    
    
    elif template_type == "TEMPLATE_13":
    
        uploaded_file.seek(0)
    
        result = transform_template_13(
            uploaded_file
        )
    
        return (
            result,
            "Template 13｜橡木桶月份實銷",
        )
    
    
    elif template_type == "TEMPLATE_14":
    
        uploaded_file.seek(0)
    
        result = transform_template_14(
            uploaded_file
        )
    
        return (
            result,
            "Template 14｜美滿商行商品彙總",
        )
    
    
    elif template_type == "TEMPLATE_15":
    
        uploaded_file.seek(0)
    
        result = transform_template_15(
            uploaded_file
        )
    
        return (
            result,
            "Template 15｜Multi-sheet SKU / QTY",
        )
        


    
    # ========================================================
    # UNKNOWN
    # ========================================================

    else:

        raise ValueError(
            f"無法辨識此檔案的 Template 格式：{template_type}"
        )

# ============================================================
# Mapping Helpers
# ============================================================
def resolve_required(
    df,
    required,
):
    lookup = {
        norm_key(c): c
        for c in df.columns
    }

    resolved = {}
    missing = []

    for required_name in required:
        key = norm_key(
            required_name
        )

        if key in lookup:
            resolved[
                required_name
            ] = lookup[key]
        else:
            missing.append(
                required_name
            )

    return (
        resolved,
        missing,
    )

# ============================================================
# Standard Header Mapping
#
# 不同 Template 的原始欄位名稱可能不同，
# 但最後都會統一轉成標準欄位。
# ============================================================

STANDARD_HEADER_MAPPING = {

    # --------------------------------------------------------
    # SKU / Product Name
    # 最後統一 -> RAW_SKU -> Final Report「品項」
    # --------------------------------------------------------

    "SKU": [
        "貨品名稱",
        "商品名稱",
        "品項名稱",
        "品名",
        "商品品名",
        "產品名稱",
        "SKU Name",
        "Product Name",
        "產品品名",
        "單品名稱",
    ],


    # --------------------------------------------------------
    # Quantity
    # 最後統一 -> QTY -> Final Report「銷量(瓶)」
    # --------------------------------------------------------

    "QTY": [
        "數量",
        "數 量",
        "銷量",
        "銷量 Quantity",
        "Quantity",
        "QTY",
        "瓶數",
    ],


    # --------------------------------------------------------
    # Product Code
    # --------------------------------------------------------

    "PRODUCT_CODE": [
        "商品號碼",
        "商品代號",
        "商品編號",
        "貨品號碼",
        "貨品代號",
        "貨品編號",
        "品項代碼",
        "SKU Code",
        "SKU code",
        "產品代碼",
    ],


    # --------------------------------------------------------
    # Customer
    # --------------------------------------------------------

    "CUSTOMER": [
        "客戶名稱",
        "客戶",
        "店家名稱",
        "門市名稱",
        "Customer Name",
    ],


    # --------------------------------------------------------
    # Unit Price
    # Source price only
    # Final Report still uses RSP
    # --------------------------------------------------------

    "UNIT_PRICE": [
        "單價",
        "售價",
        "價格",
        "Unit Price",
        "Price",
    ],


    # --------------------------------------------------------
    # Total Price / Amount
    # --------------------------------------------------------

    "TOTAL_PRICE": [
        "小計",
        "小計金額",
        "總價",
        "總金額",
        "銷貨金額",
        "金額",
        "Amount",
        "Total",
        "Total Price",
    ],


    # --------------------------------------------------------
    # Brand
    # --------------------------------------------------------

    "BRAND": [
        "品牌",
        "品牌名稱",
        "品牌名稱 Brand Name",
        "Brand",
        "Brand Name",
    ],


    # --------------------------------------------------------
    # Size
    # --------------------------------------------------------

    "SIZE": [
        "容量",
        "規格",
        "規格 Size (ml)",
        "Size",
        "Size (ml)",
    ],


    # --------------------------------------------------------
    # Category
    # --------------------------------------------------------

    "CATEGORY": [
        "品類",
        "品類 Category",
        "Category",
    ],
}



# ============================================================
# Find Header Row by Standard Mapping
# ============================================================

def find_header_row(
    raw,
    required_fields,
    max_rows=30,
):

    """
    Example:

    required_fields = [
        "SKU",
        "QTY",
    ]

    系統會去 STANDARD_HEADER_MAPPING 裡面找可能的 header 名稱。

    如果同一列同時找到 SKU + QTY，
    就視為 header row。
    """

    for i in range(
        min(
            max_rows,
            len(raw),
        )
    ):

        values = [
            norm_text(value)
            for value
            in raw.iloc[i].tolist()
        ]


        all_found = True


        for field_name in required_fields:

            candidates = (
                STANDARD_HEADER_MAPPING.get(
                    field_name,
                    [],
                )
            )


            field_found = any(

                any(
                    norm_key(candidate)
                    ==
                    norm_key(value)

                    for candidate
                    in candidates
                )

                for value
                in values

                if norm_text(value)
            )


            if not field_found:

                all_found = False
                break


        if all_found:

            return i


    return None



# ============================================================
# Customer Mapping
# ============================================================
# ============================================================
# CUSTOMER MAPPING PREPARATION
# ============================================================

def prepare_customer_mapping(df):

    resolved, missing = resolve_required(
        df,
        CUSTOMER_REQUIRED,
    )

    if missing:

        raise ValueError(
            "Customer Mapping 缺少欄位："
            + ", ".join(missing)
        )


    mapping = pd.DataFrame({

        column:
            df[source]

        for column, source
        in resolved.items()
    })


    # ========================================================
    # IMPORTANT
    #
    # ID 欄位一定使用 norm_id_key
    # 不要使用 norm_key
    # ========================================================

    mapping["OUTLET_KEY"] = (
        mapping["Outlet No"]
        .map(
            norm_id_key
        )
    )


    mapping["SALES_KEY"] = (
        mapping["Sales ID"]
        .map(
            norm_id_key
        )
    )


    mapping["CONTRACT_KEY"] = (
        mapping["Contract JDE"]
        .map(
            norm_id_key
        )
    )


    # ========================================================
    # Name 欄位才使用 norm_key
    # ========================================================

    mapping["RAWDATA_KEY"] = (
        mapping["Rawdata Name"]
        .map(
            norm_key
        )
    )


    # ========================================================
    # Duplicate Outlet Check
    # ========================================================

    duplicate_outlet = mapping[
        mapping["OUTLET_KEY"].ne("")
        &
        mapping.duplicated(
            subset=[
                "OUTLET_KEY"
            ],
            keep=False,
        )
    ].copy()


    return (
        mapping,
        duplicate_outlet,
    )

# ============================================================
# SKU Mapping
# ============================================================
# ============================================================
# SKU MAPPING
#
# JOIN KEY:
# Template RAW_SKU
#      ↓
# SKU Mapping["SKU Name"]
#
# RETURN:
# Mapping Name
# RSP
# Price Band
# Manufacture
# Brand
# SIZE
# CATEGORY
# ============================================================
def prepare_sku_mapping(df):

    df = df.copy()

    # --------------------------------------------------------
    # Clean column names
    # --------------------------------------------------------

    df.columns = [
        norm_text(col)
        for col in df.columns
    ]


    # --------------------------------------------------------
    # Find required columns
    # --------------------------------------------------------

    resolved, missing = (
        resolve_required(
            df,
            SKU_REQUIRED,
        )
    )


    if missing:

        raise ValueError(
            "SKU Mapping 缺少欄位："
            + ", ".join(missing)
            + f"\n目前欄位：{list(df.columns)}"
        )


    # --------------------------------------------------------
    # Standardise Mapping DataFrame
    # --------------------------------------------------------

    mapping = pd.DataFrame({

        "Mapping Name":
            df[
                resolved[
                    "Mapping Name"
                ]
            ],

        "RSP":
            df[
                resolved[
                    "RSP"
                ]
            ],

        "Price Band":
            df[
                resolved[
                    "Price Band"
                ]
            ],

        "Manufacture":
            df[
                resolved[
                    "Manufacture"
                ]
            ],

        "Brand":
            df[
                resolved[
                    "Brand"
                ]
            ],

        "SKU Name":
            df[
                resolved[
                    "SKU Name"
                ]
            ],

        "SIZE":
            df[
                resolved[
                    "SIZE"
                ]
            ],

        "CATEGORY":
            df[
                resolved[
                    "CATEGORY"
                ]
            ],
    })


    # --------------------------------------------------------
    # IMPORTANT
    #
    # SKU Name is JOIN KEY
    # NOT Mapping Name
    # --------------------------------------------------------

    mapping[
        "SKU_KEY"
    ] = (

        mapping[
            "SKU Name"
        ]

        .map(
            norm_key
        )
    )


    # --------------------------------------------------------
    # Remove blank SKU keys
    # --------------------------------------------------------

    mapping = mapping[
        mapping[
            "SKU_KEY"
        ].ne("")
    ].copy()


    # --------------------------------------------------------
    # Duplicate SKU Name check
    # --------------------------------------------------------

    duplicate_mapping = mapping[
        mapping.duplicated(
            subset=[
                "SKU_KEY"
            ],
            keep=False,
        )
    ].copy()


    # --------------------------------------------------------
    # Mapping table used for join
    #
    # Keep first only prevents duplicate join expansion.
    # Duplicate records remain visible in duplicate report.
    # --------------------------------------------------------

    mapping_for_join = (

        mapping

        .drop_duplicates(
            subset=[
                "SKU_KEY"
            ],
            keep="first",
        )

        .reset_index(
            drop=True
        )
    )


    return (
        mapping_for_join,
        duplicate_mapping,
    )


# ============================================================
# Mapping File Reader
# ============================================================
def read_mapping_file(
    uploaded_file,
    preferred_sheet=None,
):
    suffix = Path(
        uploaded_file.name
    ).suffix.lower()

    if suffix == ".csv":
        return read_csv_flexible(
            uploaded_file,
            header=0,
        )

    uploaded_file.seek(0)

    excel_file = pd.ExcelFile(
        uploaded_file
    )

    if (
        preferred_sheet
        and preferred_sheet in excel_file.sheet_names
    ):
        sheet = preferred_sheet
    else:
        sheet = excel_file.sheet_names[0]

    uploaded_file.seek(0)

    return pd.read_excel(
        uploaded_file,
        sheet_name=sheet,
        header=0,
    )


# ============================================================
# Customer Mapping Join
#
# Priority:
# 1. Outlet No
# 2. Rawdata Name fallback
#
# LEFT JOIN only
# ============================================================
# ============================================================
# CUSTOMER MAPPING JOIN
#
# Priority:
# 1. Outlet No
# 2. Rawdata Name fallback
#
# LEFT JOIN only
# ============================================================

# ============================================================
# CUSTOMER MAPPING JOIN
#
# Priority:
# 1. Outlet No
# 2. Rawdata Name
#
# IMPORTANT:
# ID -> norm_id_key
# Name -> norm_key
# ============================================================

def attach_customer_mapping(
    raw,
    customer_map,
):

    source = raw.copy()


    # ========================================================
    # 1. SOURCE KEYS
    # ========================================================

    source["FILE_OUTLET_KEY"] = (
        source["FILE_OUTLET_NO"]
        .map(
            norm_id_key
        )
    )


    source["FILE_RAWDATA_KEY"] = (
        source["FILE_RAWDATA_NAME"]
        .map(
            norm_key
        )
    )


    # ========================================================
    # 2. FIRST MATCH BY OUTLET NO
    # ========================================================

    by_outlet = (
        customer_map[
            customer_map[
                "OUTLET_KEY"
            ].ne("")
        ]
        .drop_duplicates(
            subset=[
                "OUTLET_KEY"
            ],
            keep="first",
        )
        .copy()
    )


    merged = source.merge(
        by_outlet,
        how="left",
        left_on="FILE_OUTLET_KEY",
        right_on="OUTLET_KEY",
        suffixes=(
            "",
            "_CM",
        ),
    )


    # ========================================================
    # 3. FIND UNMATCHED
    # ========================================================

    unmatched = (
        merged[
            "Outlet No"
        ].isna()
        |
        merged[
            "Outlet No"
        ]
        .map(
            norm_text
        )
        .eq("")
    )


    # ========================================================
    # 4. FALLBACK BY RAWDATA NAME
    # ========================================================

    if unmatched.any():

        by_raw = (
            customer_map[
                customer_map[
                    "RAWDATA_KEY"
                ].ne("")
            ]
            .drop_duplicates(
                subset=[
                    "RAWDATA_KEY"
                ],
                keep="first",
            )
            .copy()
        )


        fallback_source = (
            source.loc[
                unmatched.values
            ]
            .copy()
        )


        fallback = fallback_source.merge(
            by_raw,
            how="left",
            left_on="FILE_RAWDATA_KEY",
            right_on="RAWDATA_KEY",
            suffixes=(
                "",
                "_CM",
            ),
        )


        mapping_columns = (
            CUSTOMER_REQUIRED
            + [
                "OUTLET_KEY",
                "RAWDATA_KEY",
                "SALES_KEY",
                "CONTRACT_KEY",
            ]
        )


        for column in mapping_columns:

            if (
                column in fallback.columns
                and
                column in merged.columns
            ):

                merged.loc[
                    unmatched,
                    column
                ] = fallback[
                    column
                ].values


    # ========================================================
    # 5. MAPPING STATUS
    # ========================================================

    merged[
        "CUSTOMER_MAPPING_STATUS"
    ] = (
        merged[
            "Outlet No"
        ]
        .apply(
            lambda value:
                "MAPPED"
                if norm_text(value)
                else "UNMAPPED"
        )
    )


    # ========================================================
    # 6. SALES ID CHECK
    # ========================================================

    merged[
        "SALES_ID_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if (
                row[
                    "CUSTOMER_MAPPING_STATUS"
                ]
                == "UNMAPPED"
            )
            else (
                "MATCH"
                if (
                    norm_id_key(
                        row[
                            "FILE_SALES_ID"
                        ]
                    )
                    ==
                    norm_id_key(
                        row[
                            "Sales ID"
                        ]
                    )
                )
                else "MISMATCH"
            )
        ),
        axis=1,
    )


    # ========================================================
    # 7. CONTRACT CHECK
    # ========================================================

    merged[
        "CONTRACT_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if (
                row[
                    "CUSTOMER_MAPPING_STATUS"
                ]
                == "UNMAPPED"
            )
            else (
                "MATCH"
                if (
                    norm_id_key(
                        row[
                            "FILE_CONTRACT_JDE"
                        ]
                    )
                    ==
                    norm_id_key(
                        row[
                            "Contract JDE"
                        ]
                    )
                )
                else "MISMATCH"
            )
        ),
        axis=1,
    )


    # ========================================================
    # 8. OUTLET CHECK
    # ========================================================

    merged[
        "OUTLET_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if (
                row[
                    "CUSTOMER_MAPPING_STATUS"
                ]
                == "UNMAPPED"
            )
            else (
                "MATCH"
                if (
                    norm_id_key(
                        row[
                            "FILE_OUTLET_NO"
                        ]
                    )
                    ==
                    norm_id_key(
                        row[
                            "Outlet No"
                        ]
                    )
                )
                else "MISMATCH"
            )
        ),
        axis=1,
    )


    # ========================================================
    # 9. RAWDATA NAME CHECK
    # ========================================================

    merged[
        "RAWDATA_NAME_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if (
                row[
                    "CUSTOMER_MAPPING_STATUS"
                ]
                == "UNMAPPED"
            )
            else (
                "MATCH"
                if (
                    norm_key(
                        row[
                            "FILE_RAWDATA_NAME"
                        ]
                    )
                    ==
                    norm_key(
                        row[
                            "Rawdata Name"
                        ]
                    )
                )
                else "MISMATCH"
            )
        ),
        axis=1,
    )


    return merged

    
# ============================================================
# SKU Mapping Join
# ============================================================
# ============================================================
# SKU MAPPING JOIN
# ============================================================

def attach_sku_mapping(
    df,
    sku_map,
):

    output = df.copy()


    # --------------------------------------------------------
    # Template 商品名稱
    # RAW_SKU -> normalized SKU_KEY
    # --------------------------------------------------------

    output[
        "SKU_KEY"
    ] = (

        output[
            "RAW_SKU"
        ]

        .map(
            norm_key
        )
    )


    # --------------------------------------------------------
    # LEFT JOIN
    #
    # Template RAW_SKU
    #          ↓
    # SKU_KEY
    #          =
    # Mapping SKU Name
    #
    # LEFT JOIN = unmapped source rows remain
    # --------------------------------------------------------

    output = output.merge(

        sku_map[
            [
                "SKU_KEY",
                "SKU Name",
                "Mapping Name",
                "RSP",
                "Price Band",
                "Manufacture",
                "Brand",
                "SIZE",
                "CATEGORY",
            ]
        ],

        how="left",

        on="SKU_KEY",
    )


    # --------------------------------------------------------
    # Mapping Status
    # --------------------------------------------------------

    output[
        "SKU_MAPPING_STATUS"
    ] = (

        output[
            "Mapping Name"
        ]

        .apply(

            lambda value:

            "MAPPED"

            if norm_text(value)

            else "UNMAPPED"
        )
    )


    return output


# ============================================================
# Build Final Report
# ============================================================
# ============================================================
# SAFE COLUMN GETTER
# ============================================================

def safe_column(
    df,
    column_name,
    default=pd.NA,
):

    if column_name in df.columns:

        return df[
            column_name
        ]

    return pd.Series(
        [default] * len(df),
        index=df.index,
    )


# ============================================================
# BUILD FINAL REPORT
#
# PRICE LOGIC:
# 單價     = RSP
# 總價     = QTY * RSP
# 建議售價 = RSP
#
# SKU LOGIC:
# Template 商品名稱
#       ↓
# RAW_SKU
#       ↓
# JOIN SKU Mapping["SKU Name"]
#       ↓
# Mapping Name / Brand / RSP / etc.
# ============================================================

def build_report(detail):

    report = pd.DataFrame(
        index=detail.index
    )


    # ========================================================
    # 1. CUSTOMER MAPPING FIELDS
    # ========================================================

    report["地區"] = safe_column(
        detail,
        "REGION",
    )


    report["業務員"] = safe_column(
        detail,
        "Sales",
    )


    report["合約等級"] = safe_column(
        detail,
        "CONTRACT TYPE",
    )


    report["合約編號"] = safe_column(
        detail,
        "Contract JDE",
    )


    report["合約名稱"] = safe_column(
        detail,
        "Contract NAME",
    )


    report["店家編號"] = safe_column(
        detail,
        "FILE_OUTLET_NO",
    )


    report["店家名稱"] = safe_column(
        detail,
        "Outlet NAME",
    )


    # ========================================================
    # 2. SALES QUANTITY
    # ========================================================

    qty = pd.to_numeric(
        safe_column(
            detail,
            "QTY",
        ),
        errors="coerce",
    )


    report["銷量(瓶)"] = qty


    # ========================================================
    # 3. RSP
    #
    # All price-related fields use SKU Mapping RSP
    # ========================================================

    rsp = pd.to_numeric(
        safe_column(
            detail,
            "RSP",
        ),
        errors="coerce",
    )


    # 單價 = RSP
    report["單價"] = rsp


    # 總價 = QTY × RSP
    report["總價"] = (
        qty
        *
        rsp
    )


    # 建議售價 = RSP
    report["建議售價"] = rsp


    # ========================================================
    # 4. SKU MAPPING FIELDS
    # ========================================================

    report["價格帶"] = safe_column(
        detail,
        "Price Band",
    )


    report["製造商"] = safe_column(
        detail,
        "Manufacture",
    )


    report["品牌"] = safe_column(
        detail,
        "Brand",
    )


    report["統一品項名稱"] = safe_column(
        detail,
        "Mapping Name",
    )


    # ========================================================
    # 5. ORIGINAL SKU NAME
    #
    # 保留原始 Template 商品名稱
    # ========================================================

    report["品項"] = safe_column(
        detail,
        "RAW_SKU",
    )


    report["容量"] = safe_column(
        detail,
        "SIZE",
    )


    report["品類"] = safe_column(
        detail,
        "CATEGORY",
    )


    # ========================================================
    # 6. YEAR / MONTH
    # ========================================================

    report["年"] = pd.to_numeric(
        safe_column(
            detail,
            "YEAR",
        ),
        errors="coerce",
    ).astype(
        "Int64"
    )


    report["月"] = pd.to_numeric(
        safe_column(
            detail,
            "MONTH",
        ),
        errors="coerce",
    ).astype(
        "Int64"
    )


    # ========================================================
    # 7. ENSURE ALL OUTPUT COLUMNS EXIST
    # ========================================================

    for column in OUTPUT_COLUMNS:

        if column not in report.columns:

            report[column] = pd.NA


    # ========================================================
    # 8. FORCE FINAL COLUMN ORDER
    # ========================================================

    report = report[
        OUTPUT_COLUMNS
    ]


    return report

# ============================================================
# Excel Export
# ============================================================
def excel_bytes(sheets):
    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl",
    ) as writer:

        for (
            sheet_name,
            dataframe
        ) in sheets.items():

            dataframe.to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False,
            )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# UI - Template Upload
# ============================================================
# ============================================================
# STEP 1 - Upload Monthly Template
# ============================================================

section_header(
    "1",
    "上傳當月份資料",
    "Upload monthly POS templates",
)

template_files = st.file_uploader(
    "可一次上傳多個 Excel / CSV Template",
    type=[
        "xlsx",
        "xls",
        "csv",
    ],
    accept_multiple_files=True,
    key="templates",
)

st.caption(
    "系統會自動辨識不同 Template 格式，不需要人工選擇格式。"
)

# ============================================================
# STEP 2 - Upload Mapping
# ============================================================

section_header(
    "2",
    "上傳 Mapping 資料",
    "Upload customer and SKU mapping files",
)

map_col1, map_col2 = st.columns(2)

with map_col1:

    customer_file = st.file_uploader(
        "Customer Mapping",
        type=[
            "xlsx",
            "xls",
            "csv",
        ],
        key="customer",
    )

with map_col2:

    sku_files = st.file_uploader(
        "SKU Mapping",
        type=[
            "xlsx",
            "xls",
            "csv",
        ],
        accept_multiple_files=True,
        key="sku",
    )

    if sku_files:

        st.caption(
            f"已上傳 {len(sku_files)} 個 SKU Mapping 檔案"
        )


# ============================================================
# STEP 3 - DASHBOARD
# ============================================================

section_header(
    "3",
    "Monthly Dashboard",
    "Submission status, mapping quality and data reconciliation",
)




# ============================================================
# Basic Validation
# ============================================================
if not template_files:
    st.info(
        "請至少上傳一個當月份 Template。"
    )
    st.stop()

if (
    customer_file is None
    or not sku_files
):

    st.info(
        "請上傳 Customer Mapping 與至少一個 SKU Mapping。"
    )

    st.stop()


# ============================================================
# Load Customer Mapping
# ============================================================
try:
    customer_raw = read_mapping_file(
        customer_file,
        preferred_sheet="Customer data",
    )

    (
        customer_map,
        customer_duplicates,
    ) = prepare_customer_mapping(
        customer_raw
    )

except Exception as e:
    st.error(
        f"Customer Mapping 讀取失敗：{e}"
    )
    st.stop()


# ============================================================
# ============================================================
# Load Multiple SKU Mapping Files
# ============================================================

try:

    sku_frames = []
    sku_file_logs = []


    # ========================================================
    # Read Each SKU Mapping File
    # ========================================================

    for sku_file in sku_files:

        try:

            sku_file.seek(0)

            sku_temp = read_mapping_file(
                sku_file,
                preferred_sheet="SKU data",
            )


            # =================================================
            # Validate Read Result
            # =================================================

            if sku_temp is None:

                raise ValueError(
                    "讀取結果為 None"
                )


            if not isinstance(
                sku_temp,
                pd.DataFrame,
            ):

                raise ValueError(
                    "讀取結果不是 DataFrame"
                )


            if sku_temp.empty:

                raise ValueError(
                    "SKU Mapping 沒有資料"
                )


            # =================================================
            # Copy
            # =================================================

            sku_temp = (
                sku_temp
                .copy()
            )


            # =================================================
            # Record Source File
            #
            # 未來 Duplicate 可以知道來源檔案
            # =================================================

            sku_temp[
                "SKU_MAPPING_SOURCE_FILE"
            ] = sku_file.name


            # =================================================
            # Add to Merge List
            # =================================================

            sku_frames.append(
                sku_temp
            )


            # =================================================
            # Success Log
            # =================================================

            sku_file_logs.append({

                "檔案":
                    sku_file.name,

                "狀態":
                    "SUCCESS",

                "筆數":
                    len(
                        sku_temp
                    ),

                "錯誤":
                    "",

            })


        except Exception as file_error:

            # =================================================
            # Failed File
            #
            # 單一 SKU Mapping 失敗，
            # 不影響其他 Mapping File
            # =================================================

            sku_file_logs.append({

                "檔案":
                    sku_file.name,

                "狀態":
                    "FAILED",

                "筆數":
                    0,

                "錯誤":
                    (
                        f"{type(file_error).__name__}: "
                        f"{str(file_error)}"
                    ),

            })


    # ========================================================
    # Create Upload Log FIRST
    #
    # 即使後面 prepare_sku_mapping 失敗，
    # 還是可以看到每個檔案的讀取結果
    # ========================================================

    sku_upload_log = pd.DataFrame(
        sku_file_logs
    )


    # ========================================================
    # Validate
    # 至少一個 SKU Mapping File 成功
    # ========================================================

    if not sku_frames:

        st.error(
            "所有 SKU Mapping 檔案都讀取失敗。"
        )

        st.dataframe(
            sku_upload_log,
            width="stretch",
            hide_index=True,
        )

        st.stop()


    # ========================================================
    # Combine All Successful SKU Mapping Files
    # ========================================================

    sku_raw = pd.concat(
        sku_frames,
        ignore_index=True,
        sort=False,
    )


    # ========================================================
    # Prepare SKU Mapping
    #
    # 沿用原本 prepare_sku_mapping()
    # ========================================================

    (
        sku_map,
        sku_duplicates,
    ) = prepare_sku_mapping(
        sku_raw
    )


except Exception as e:

    st.error(
        f"SKU Mapping 整併 / Mapping Prepare 失敗：{e}"
    )


    # ========================================================
    # Show File-level Log
    # 即使 prepare_sku_mapping() 失敗也能看到
    # ========================================================

    if (
        "sku_upload_log"
        in locals()
    ):

        st.dataframe(
            sku_upload_log,
            width="stretch",
            hide_index=True,
        )


    st.stop()


# ============================================================
# SKU Mapping Upload Status
# ============================================================

if (
    "sku_upload_log"
    in locals()
):

    with st.expander(
        "📦 SKU Mapping 上傳狀況",
        expanded=False,
    ):

        st.dataframe(
            sku_upload_log,
            width="stretch",
            hide_index=True,
        )


        # ====================================================
        # Summary
        # ====================================================

        sku_success_count = int(
            sku_upload_log[
                "狀態"
            ]
            .eq(
                "SUCCESS"
            )
            .sum()
        )


        sku_failed_count = int(
            sku_upload_log[
                "狀態"
            ]
            .eq(
                "FAILED"
            )
            .sum()
        )


        st.caption(
            f"成功：{sku_success_count} 個 ｜ "
            f"失敗：{sku_failed_count} 個 ｜ "
            f"合併後原始資料：{len(sku_raw):,} 筆"
        )


# ============================================================
# ============================================================
# Auto Process All Templates
# ============================================================

frames = []
upload_logs = []


for file in template_files:

    meta = parse_filename(
        file.name
    )


    # ========================================================
    # 先單獨做 Template Detection
    #
    # 這樣即使後面的 transform 失敗，
    # 我們仍然知道它原本辨識成哪個 Template
    # ========================================================

    try:

        file.seek(0)

        detected_type_code = (
            detect_template_type(
                file
            )
        )

    except Exception as detect_error:

        detected_type_code = "DETECTION_ERROR"


    try:

        file.seek(0)

        (
            data,
            detected_template,
        ) = auto_transform_template(
            file
        )

        data.attrs = {}
        
        frames.append(
            data
        )


        upload_logs.append({

            "檔案":
                file.name,

            "辨識格式":
                detected_template,

            "Template Code":
                detected_type_code,

            "狀態":
                "SUCCESS",

            "錯誤":
                "",

            "原始筆數":
                len(data),

            "業務代號":
                meta[
                    "FILE_SALES_ID"
                ],

            "合約編號":
                meta[
                    "FILE_CONTRACT_JDE"
                ],

            "年月":
                meta[
                    "FILE_PERIOD"
                ],

            "店家編號":
                meta[
                    "FILE_OUTLET_NO"
                ],

            "分店/Rawdata Name":
                meta[
                    "FILE_RAWDATA_NAME"
                ],

            "檔名格式":
                meta[
                    "FILENAME_STATUS"
                ],
        })


    except Exception as e:

        upload_logs.append({

            "檔案":
                file.name,

            # 不要再強制 UNKNOWN
            "辨識格式":
                detected_type_code,

            "Template Code":
                detected_type_code,

            "狀態":
                "FAILED",

            "錯誤":
                str(e),

            "原始筆數":
                0,

            "業務代號":
                meta[
                    "FILE_SALES_ID"
                ],

            "合約編號":
                meta[
                    "FILE_CONTRACT_JDE"
                ],

            "年月":
                meta[
                    "FILE_PERIOD"
                ],

            "店家編號":
                meta[
                    "FILE_OUTLET_NO"
                ],

            "分店/Rawdata Name":
                meta[
                    "FILE_RAWDATA_NAME"
                ],

            "檔名格式":
                meta[
                    "FILENAME_STATUS"
                ],
        })


upload_log = pd.DataFrame(
    upload_logs
)

# ============================================================
# Fail Safe
# ============================================================
if not frames:
    st.error(
        "所有 Template 都處理失敗。"
    )

    st.dataframe(
        upload_log,
        use_container_width=True,
        hide_index=True,
    )

    st.stop()


# ============================================================
# Combine All Templates
# ============================================================
raw_all = pd.concat(
    frames,
    ignore_index=True,
    sort=False,
)


# ============================================================
# Customer Mapping
# ============================================================
with_customer = attach_customer_mapping(
    raw_all,
    customer_map,
)


# ============================================================
# SKU Mapping
# ============================================================
detail = attach_sku_mapping(
    with_customer,
    sku_map,
)


# ============================================================
# Final Report
# ============================================================
final_report = build_report(
    detail
)


# ============================================================
# Dashboard Metrics
# ============================================================
source_count = len(
    raw_all
)

output_count = len(
    final_report
)

customer_unmapped = int(
    detail[
        "CUSTOMER_MAPPING_STATUS"
    ]
    .eq("UNMAPPED")
    .sum()
)

sku_unmapped = int(
    detail[
        "SKU_MAPPING_STATUS"
    ]
    .eq("UNMAPPED")
    .sum()
)

both_unmapped = int(
    (
        detail[
            "CUSTOMER_MAPPING_STATUS"
        ].eq("UNMAPPED")
        &
        detail[
            "SKU_MAPPING_STATUS"
        ].eq("UNMAPPED")
    )
    .sum()
)

sales_mismatch = int(
    detail[
        "SALES_ID_CHECK"
    ]
    .eq("MISMATCH")
    .sum()
)

contract_mismatch = int(
    detail[
        "CONTRACT_CHECK"
    ]
    .eq("MISMATCH")
    .sum()
)

failed_files = int(
    upload_log[
        "狀態"
    ]
    .eq("FAILED")
    .sum()
)

successful_files = int(
    upload_log[
        "狀態"
    ]
    .eq("SUCCESS")
    .sum()
)



# ============================================================
# KPI ROW
# ============================================================

d1, d2, d3 = st.columns(3)
d4, d5, d6 = st.columns(3)


with d1:

    st.metric(
        "🐄 來源資料",
        f"{source_count:,}",
    )


with d2:

    st.metric(
        "🧀 輸出資料",
        f"{output_count:,}",
        delta=(
            f"{output_count - source_count:+,}"
        ),
    )


with d3:

    st.metric(
        "🍪 Customer Unmapped",
        f"{customer_unmapped:,}",
    )


with d4:

    st.metric(
        "🍾 SKU Unmapped",
        f"{sku_unmapped:,}",
    )


with d5:

    st.metric(
        "✔️ 成功檔案",
        f"{successful_files:,}",
    )


with d6:

    st.metric(
        "❌ 失敗檔案",
        f"{failed_files:,}",
    )

# ============================================================
# Reconciliation
# ============================================================
if (
    source_count
    == output_count
):
    st.success(
        f"✔️ 資料筆數一致：來源 {source_count:,} = 輸出 {output_count:,}。"
        " Unmapped 資料仍保留，不會因 Mapping 失敗被刪除。"
    )

else:
    st.error(
        f"❌ 資料筆數異常：來源 {source_count:,}，"
        f"輸出 {output_count:,}，"
        f"差異 {output_count - source_count:+,}。"
    )


# ============================================================
# Additional Quality Metrics
# ============================================================
q1, q2, q3 = st.columns(3)

q1.metric(
    "Customer + SKU 同時 Unmapped",
    f"{both_unmapped:,}",
)

q2.metric(
    "業務代號不一致",
    f"{sales_mismatch:,}",
)

q3.metric(
    "合約編號不一致",
    f"{contract_mismatch:,}",
)


# ============================================================
# ============================================================
# MONTHLY SUBMISSION TRACKING
# ============================================================


# ============================================================
# 1. NORMALIZE PERIOD
#
# 202609 -> 202609
# 11509  -> 202609
# ============================================================

def normalize_period_key(value):

    text = norm_text(value)

    digits = re.sub(
        r"\D",
        "",
        text,
    )

    # YYYYMM
    if len(digits) == 6:

        year = int(
            digits[:4]
        )

        month = int(
            digits[4:]
        )

        if (
            1900 <= year <= 2200
            and
            1 <= month <= 12
        ):
            return f"{year}{month:02d}"


    # ROC YYYMM
    # 11509 -> 202609
    if len(digits) == 5:

        year = (
            int(digits[:3])
            + 1911
        )

        month = int(
            digits[3:]
        )

        if 1 <= month <= 12:
            return f"{year}{month:02d}"


    return ""


# ============================================================
# 2. PREPARE ACTUAL UPLOAD
# ============================================================

actual_upload = (
    upload_log
    .copy()
)


# ============================================================
# 3. PERIOD KEY
# ============================================================

actual_upload[
    "PERIOD_KEY"
] = (
    actual_upload[
        "年月"
    ]
    .map(
        normalize_period_key
    )
)


# ============================================================
# 4. AVAILABLE MONTHS
# ============================================================

available_periods = sorted(
    actual_upload[
        "PERIOD_KEY"
    ]
    .replace(
        "",
        pd.NA,
    )
    .dropna()
    .unique(),
    reverse=True,
)


# ============================================================
# 5. MONTH SELECTOR
# ============================================================

if available_periods:

    selected_period = st.selectbox(
        "📅 選擇繳交月份",
        options=available_periods,
        index=0,
        format_func=lambda x:
            f"{x[:4]} / {x[4:]}",
        key="reporting_period_selector",
    )

else:

    selected_period = ""


reporting_period = selected_period


# ============================================================
# 6. EXPECTED SUBMISSION LIST
#
# Customer Mapping = 每月應繳名單
# ============================================================

expected_submission = (
    customer_map[
        [
            "Sales ID",
            "Sales",
            "Contract JDE",
            "Contract NAME",
            "Outlet No",
            "Outlet NAME",
            "Rawdata Name",
        ]
    ]
    .copy()
)


# ============================================================
# 7. REMOVE INVALID OUTLET
# ============================================================

expected_submission = (
    expected_submission[
        expected_submission[
            "Outlet No"
        ]
        .map(
            norm_text
        )
        .ne("")
    ]
    .copy()
)


# ============================================================
# 8. EXPECTED KEYS
# ============================================================

expected_submission[
    "PERIOD_KEY"
] = selected_period


expected_submission[
    "SALES_KEY"
] = (
    expected_submission[
        "Sales ID"
    ]
    .map(
        norm_id_key
    )
)


expected_submission[
    "CONTRACT_KEY"
] = (
    expected_submission[
        "Contract JDE"
    ]
    .map(
        norm_id_key
    )
)


expected_submission[
    "OUTLET_TRACKING_KEY"
] = (
    expected_submission[
        "Outlet No"
    ]
    .map(
        norm_id_key
    )
)


# ============================================================
# 9. EXPECTED KEY
#
# YYYYMM | Sales | Contract | Outlet
# ============================================================

expected_submission[
    "EXPECTED_KEY"
] = (
    expected_submission[
        "PERIOD_KEY"
    ]
    + "|"
    + expected_submission[
        "SALES_KEY"
    ]
    + "|"
    + expected_submission[
        "CONTRACT_KEY"
    ]
    + "|"
    + expected_submission[
        "OUTLET_TRACKING_KEY"
    ]
)


# ============================================================
# 10. EXPECTED DUPLICATE REMOVAL
# ============================================================

expected_submission = (
    expected_submission
    .drop_duplicates(
        subset=[
            "EXPECTED_KEY"
        ],
        keep="first",
    )
    .reset_index(
        drop=True
    )
)


# ============================================================
# 11. ACTUAL UPLOAD KEYS
# ============================================================

actual_upload[
    "SALES_KEY"
] = (
    actual_upload[
        "業務代號"
    ]
    .map(
        norm_id_key
    )
)


actual_upload[
    "CONTRACT_KEY"
] = (
    actual_upload[
        "合約編號"
    ]
    .map(
        norm_id_key
    )
)


actual_upload[
    "OUTLET_TRACKING_KEY"
] = (
    actual_upload[
        "店家編號"
    ]
    .map(
        norm_id_key
    )
)


actual_upload[
    "UPLOAD_KEY"
] = (
    actual_upload[
        "PERIOD_KEY"
    ]
    + "|"
    + actual_upload[
        "SALES_KEY"
    ]
    + "|"
    + actual_upload[
        "CONTRACT_KEY"
    ]
    + "|"
    + actual_upload[
        "OUTLET_TRACKING_KEY"
    ]
)


# ============================================================
# 12. FILTER SELECTED MONTH
# ============================================================

actual_upload_month = (
    actual_upload[
        actual_upload[
            "PERIOD_KEY"
        ]
        == selected_period
    ]
    .copy()
)


# ============================================================
# 13. STATUS PRIORITY
#
# SUCCESS > FAILED
# ============================================================

actual_upload_month[
    "_STATUS_PRIORITY"
] = (
    actual_upload_month[
        "狀態"
    ]
    .map({
        "SUCCESS": 2,
        "FAILED": 1,
    })
    .fillna(0)
    .astype(int)
)


# ============================================================
# 14. DUPLICATE UPLOAD PROCESSING
#
# 同月份、同業務、同合約、同店家
# 如果有 SUCCESS，就保留 SUCCESS
# ============================================================

if actual_upload_month.empty:

    actual_upload_best = (
        actual_upload_month
        .drop(
            columns=[
                "_STATUS_PRIORITY"
            ],
            errors="ignore",
        )
        .copy()
    )

else:

    actual_upload_best = (
        actual_upload_month
        .sort_values(
            by=[
                "UPLOAD_KEY",
                "_STATUS_PRIORITY",
            ],
            ascending=[
                True,
                False,
            ],
            na_position="last",
        )
        .drop_duplicates(
            subset=[
                "UPLOAD_KEY"
            ],
            keep="first",
        )
        .drop(
            columns=[
                "_STATUS_PRIORITY"
            ],
            errors="ignore",
        )
        .reset_index(
            drop=True
        )
    )


# ============================================================
# 15. REQUIRED UPLOAD COLUMNS
# ============================================================

required_upload_columns = [
    "UPLOAD_KEY",
    "PERIOD_KEY",
    "檔案",
    "辨識格式",
    "狀態",
    "錯誤",
    "原始筆數",
    "年月",
    "檔名格式",
    "業務代號",
    "合約編號",
    "店家編號",
    "分店/Rawdata Name",
]


for column in required_upload_columns:

    if column not in actual_upload_best.columns:

        actual_upload_best[
            column
        ] = pd.NA


# ============================================================
# 16. EXPECTED VS ACTUAL
# ============================================================

submission_detail = (
    expected_submission
    .merge(
        actual_upload_best[
            required_upload_columns
        ],
        how="left",
        left_on="EXPECTED_KEY",
        right_on="UPLOAD_KEY",
        suffixes=(
            "",
            "_UPLOAD",
        ),
    )
)


# ============================================================
# 17. SUBMISSION STATUS
# ============================================================

def get_submission_status(row):

    if pd.isna(
        row.get(
            "檔案"
        )
    ):
        return "MISSING"

    if (
        row.get(
            "狀態"
        )
        == "SUCCESS"
    ):
        return "SUCCESS"

    return "FAILED"


submission_detail[
    "SUBMISSION_STATUS"
] = (
    submission_detail
    .apply(
        get_submission_status,
        axis=1,
    )
)


# ============================================================
# 18. UNMATCHED UPLOAD
#
# 有上傳，但是找不到 Customer Mapping 對應
# ============================================================

matched_upload_keys = set(
    submission_detail[
        "UPLOAD_KEY"
    ]
    .dropna()
    .astype(str)
)


unmatched_uploads = (
    actual_upload_best[
        ~actual_upload_best[
            "UPLOAD_KEY"
        ]
        .astype(str)
        .isin(
            matched_upload_keys
        )
    ]
    .copy()
)


unmatched_upload_count = len(
    unmatched_uploads
)


# ============================================================
# 19. KPI COUNTS
# ============================================================

expected_count = len(
    expected_submission
)


success_count = int(
    submission_detail[
        "SUBMISSION_STATUS"
    ]
    .eq(
        "SUCCESS"
    )
    .sum()
)


submission_failed_count = int(
    submission_detail[
        "SUBMISSION_STATUS"
    ]
    .eq(
        "FAILED"
    )
    .sum()
)


missing_count = int(
    submission_detail[
        "SUBMISSION_STATUS"
    ]
    .eq(
        "MISSING"
    )
    .sum()
)


# ============================================================
# 20. OVERALL COMPLETION
# ============================================================

if expected_count > 0:

    overall_completion = (
        success_count
        /
        expected_count
    )

else:

    overall_completion = 0


# ============================================================
# MONTHLY SUBMISSION DASHBOARD
# ============================================================

st.markdown("---")


if selected_period:

    st.markdown(
        f"### 🔄 {selected_period[:4]} / "
        f"{selected_period[4:]} 繳交進度"
    )

else:

    st.markdown(
        "### 🔄 繳交進度"
    )


# ============================================================
# 21. SUBMISSION KPI
# ============================================================

s1, s2, s3, s4, s5 = (
    st.columns(5)
)


with s1:

    st.metric(
        "📋 本月應繳 (待確認FY27合約數!)",
        f"{expected_count:,}",
    )


with s2:

    st.metric(
        "✔️ 成功繳交",
        f"{success_count:,}",
    )


with s3:

    st.metric(
        "❌ 上傳失敗",
        f"{submission_failed_count:,}",
    )


with s4:

    st.metric(
        "⏳ 尚未繳交",
        f"{missing_count:,}",
    )


with s5:

    st.metric(
        "⚠️ 無法識別",
        f"{unmatched_upload_count:,}",
    )


# ============================================================
# 22. OVERALL PROGRESS
# ============================================================

st.write(
    "**Overall Submission Progress**"
)


st.progress(
    max(
        0.0,
        min(
            1.0,
            overall_completion,
        ),
    )
)


st.caption(
    f"{success_count:,} / "
    f"{expected_count:,} "
    f"({overall_completion:.1%})"
)


# ============================================================
# 23. SALES SUBMISSION TABLE
# ============================================================

sales_submission = (
    submission_detail
    .groupby(
        [
            "Sales ID",
            "Sales",
        ],
        dropna=False,
        as_index=False,
    )
    .agg(

        應繳=(
            "EXPECTED_KEY",
            "count",
        ),

        成功=(
            "SUBMISSION_STATUS",
            lambda x:
                int(
                    (
                        x == "SUCCESS"
                    )
                    .sum()
                ),
        ),

        失敗=(
            "SUBMISSION_STATUS",
            lambda x:
                int(
                    (
                        x == "FAILED"
                    )
                    .sum()
                ),
        ),

        未繳=(
            "SUBMISSION_STATUS",
            lambda x:
                int(
                    (
                        x == "MISSING"
                    )
                    .sum()
                ),
        ),
    )
)


# ============================================================
# 24. COMPLETION RATE
# ============================================================

sales_submission[
    "完成率_RAW"
] = (
    sales_submission[
        "成功"
    ]
    /
    sales_submission[
        "應繳"
    ]
)


sales_submission[
    "完成率_RAW"
] = (
    sales_submission[
        "完成率_RAW"
    ]
    .fillna(0)
)


sales_submission[
    "完成率"
] = (
    sales_submission[
        "完成率_RAW"
    ]
    .apply(
        lambda x:
            f"{x:.1%}"
    )
)


# ============================================================
# 25. SORT SALES
# ============================================================

if not sales_submission.empty:

    sales_submission = (
        sales_submission
        .sort_values(
            by=[
                "完成率_RAW",
                "未繳",
            ],
            ascending=[
                True,
                False,
            ],
            na_position="last",
        )
        .reset_index(
            drop=True
        )
    )


# ============================================================
# 26. DISPLAY SALES SUBMISSION
# ============================================================

st.markdown(
    "####  👀業務繳交進度"
)


st.dataframe(
    sales_submission[
        [
            "Sales ID",
            "Sales",
            "應繳",
            "成功",
            "失敗",
            "未繳",
            "完成率",
        ]
    ],
    use_container_width=True,
    hide_index=True,
)



# ============================================================
# 28. SHOP DETAIL
# ============================================================

st.markdown(
    "#### 🍩 店家繳交明細"
)


shop_detail = (
    submission_detail[
        [
            "Sales ID",
            "Sales",
            "Contract JDE",
            "Contract NAME",
            "Outlet No",
            "Outlet NAME",
            "Rawdata Name",
            "檔案",
            "辨識格式",
            "SUBMISSION_STATUS",
            "錯誤",
        ]
    ]
    .copy()
)


shop_detail.insert(
    0,
    "年月",
    selected_period,
)


shop_detail = (
    shop_detail
    .rename(
        columns={

            "Sales ID":
                "業務代號",

            "Sales":
                "業務員",

            "Contract JDE":
                "合約編號",

            "Contract NAME":
                "合約名稱",

            "Outlet No":
                "店家編號",

            "Outlet NAME":
                "店家名稱",

            "SUBMISSION_STATUS":
                "繳交狀態",
        }
    )
)


# ============================================================
# 29. STATUS FILTER
# ============================================================

status_filter = st.multiselect(
    "篩選繳交狀態",
    options=[
        "SUCCESS",
        "FAILED",
        "MISSING",
    ],
    default=[
        "SUCCESS",
        "FAILED",
        "MISSING",
    ],
    key="submission_status_filter",
)


filtered_shop_detail = (
    shop_detail[
        shop_detail[
            "繳交狀態"
        ]
        .isin(
            status_filter
        )
    ]
)


st.dataframe(
    filtered_shop_detail,
    use_container_width=True,
    hide_index=True,
    height=450,
)


# ============================================================
# 30. UNMATCHED UPLOAD
# ============================================================

if unmatched_upload_count > 0:

    st.warning(
        f"⚠️ 有 {unmatched_upload_count:,} 個上傳檔案"
        "無法與 Customer Mapping 對應。"
    )


    with st.expander(
        "查看無法識別的上傳檔案"
    ):

        unmatched_columns = [
            "年月",
            "業務代號",
            "合約編號",
            "店家編號",
            "分店/Rawdata Name",
            "檔案",
            "辨識格式",
            "狀態",
            "錯誤",
            "UPLOAD_KEY",
        ]


        unmatched_columns = [
            column
            for column
            in unmatched_columns
            if column
            in unmatched_uploads.columns
        ]


        st.dataframe(
            unmatched_uploads[
                unmatched_columns
            ],
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# 31. FILE UPLOAD STATUS
# ============================================================

st.markdown(
    "#### 📁 檔案上傳狀況"
)


upload_health_columns = [
    "檔案",
    "辨識格式",
    "狀態",
    "原始筆數",
    "業務代號",
    "合約編號",
    "年月",
    "店家編號",
    "分店/Rawdata Name",
    "檔名格式",
    "錯誤",
]


upload_health_columns = [
    column
    for column
    in upload_health_columns
    if column
    in upload_log.columns
]


upload_health = (
    upload_log[
        upload_health_columns
    ]
    .copy()
)


st.dataframe(
    upload_health,
    use_container_width=True,
    hide_index=True,
    height=400,
)


# ============================================================
# 32. FILE FAILURE WARNING
# ============================================================

failed_uploads = (
    upload_log[
        upload_log[
            "狀態"
        ]
        == "FAILED"
    ]
)


if len(
    failed_uploads
) > 0:

    st.error(
        f"⚠️ 有 {len(failed_uploads):,} 個檔案處理失敗，"
        "請至「檔案上傳狀況」查看錯誤訊息。"
    )


# ============================================================
# 33. DATA QUALITY EXCEPTIONS
# ============================================================

st.markdown(
    "#### ⚠️ Data Quality Exceptions"
)


filename_error_count = int(
    (
        upload_log[
            "檔名格式"
        ]
        != "OK"
    )
    .sum()
)


e1, e2, e3, e4 = (
    st.columns(4)
)


with e1:

    st.metric(
        "🍾 SKU Unmapped",
        f"{sku_unmapped:,}",
    )


with e2:

    st.metric(
        "🍪 Customer Unmapped",
        f"{customer_unmapped:,}",
    )


with e3:

    st.metric(
        "📄 檔名格式異常",
        f"{filename_error_count:,}",
    )


with e4:

    st.metric(
        "⚠️ Submission Unmatched",
        f"{unmatched_upload_count:,}",
    )


# ============================================================
# DEBUG
#
# 測試階段先保留
# ============================================================

with st.expander(
    "🔧 Submission Tracking Debug"
):

    st.write(
        "Selected Period:",
        selected_period,
    )


    st.write(
        "Expected Submission"
    )


    st.dataframe(
        expected_submission[
            [
                "Sales ID",
                "Contract JDE",
                "Outlet No",
                "PERIOD_KEY",
                "EXPECTED_KEY",
            ]
        ],
        use_container_width=True,
        hide_index=True,
        height=250,
    )


    st.write(
        "Actual Upload"
    )


    actual_debug_columns = [
        "業務代號",
        "合約編號",
        "店家編號",
        "PERIOD_KEY",
        "UPLOAD_KEY",
        "狀態",
        "檔案",
    ]


    actual_debug_columns = [
        column
        for column
        in actual_debug_columns
        if column
        in actual_upload_best.columns
    ]


    st.dataframe(
        actual_upload_best[
            actual_debug_columns
        ],
        use_container_width=True,
        hide_index=True,
        height=250,
    )



# ============================================================
# AUDIT COLUMNS
#
# 用於 Preview / Unmapped / Audit 顯示
# ============================================================

audit_columns = [
    "SOURCE_FILE",
    "SOURCE_ROW",
    "TEMPLATE_TYPE",

    "FILE_SALES_ID",
    "FILE_CONTRACT_JDE",
    "FILE_PERIOD",
    "FILE_OUTLET_NO",
    "FILE_RAWDATA_NAME",

    "RAW_CUSTOMER",
    "RAW_ROW_CUSTOMER",

    "RAW_SKU",
    "RAW_PRODUCT_CODE",

    "QTY",

    "CUSTOMER_MAPPING_STATUS",
    "SKU_MAPPING_STATUS",

    "SALES_ID_CHECK",
    "CONTRACT_CHECK",
    "OUTLET_CHECK",
    "RAWDATA_NAME_CHECK",
]


# ============================================================
# 只保留 detail 裡實際存在的欄位
#
# 避免某些 Template 沒有特定欄位時出現 KeyError
# ============================================================

audit_columns = [
    column
    for column
    in audit_columns
    if column in detail.columns
]


# ============================================================
# STEP 4 - DATA PREVIEW
# ============================================================

section_header(
    "4",
    "Data Preview",
    "Review final output, unmapped records and mapping exceptions",
)
tabs = st.tabs([
    "最終報表",
    "Customer Unmapped",
    "SKU Unmapped",
    "全部 Unmapped",
    "上傳明細",
   # "Customer Mapping 重複",
   # "SKU Mapping 重複",
])


# Final Report
with tabs[0]:
    st.dataframe(
        final_report,
        use_container_width=True,
        hide_index=True,
        height=550,
    )


# Customer Unmapped
with tabs[1]:
    mask = (
        detail[
            "CUSTOMER_MAPPING_STATUS"
        ]
        .eq("UNMAPPED")
    )

    preview = pd.concat(
        [
            detail.loc[
                mask,
                audit_columns,
            ].reset_index(
                drop=True
            ),

            final_report.loc[
                mask
            ].reset_index(
                drop=True
            ),
        ],
        axis=1,
    )

    st.write(
        f"共 {len(preview):,} 筆"
    )

    st.dataframe(
        preview,
        use_container_width=True,
        hide_index=True,
        height=520,
    )


# SKU Unmapped
with tabs[2]:
    mask = (
        detail[
            "SKU_MAPPING_STATUS"
        ]
        .eq("UNMAPPED")
    )

    preview = pd.concat(
        [
            detail.loc[
                mask,
                audit_columns,
            ].reset_index(
                drop=True
            ),

            final_report.loc[
                mask
            ].reset_index(
                drop=True
            ),
        ],
        axis=1,
    )

    st.write(
        f"共 {len(preview):,} 筆"
    )

    st.dataframe(
        preview,
        use_container_width=True,
        hide_index=True,
        height=520,
    )


# All Unmapped
with tabs[3]:
    mask = (
        detail[
            "CUSTOMER_MAPPING_STATUS"
        ].eq("UNMAPPED")
        |
        detail[
            "SKU_MAPPING_STATUS"
        ].eq("UNMAPPED")
    )

    preview = pd.concat(
        [
            detail.loc[
                mask,
                audit_columns,
            ].reset_index(
                drop=True
            ),

            final_report.loc[
                mask
            ].reset_index(
                drop=True
            ),
        ],
        axis=1,
    )

    st.write(
        f"共 {len(preview):,} 筆"
    )

    st.dataframe(
        preview,
        use_container_width=True,
        hide_index=True,
        height=520,
    )


# Upload Log
with tabs[4]:
    st.dataframe(
        upload_log,
        use_container_width=True,
        hide_index=True,
        height=520,
    )


# Customer Duplicate Mapping
with tabs[5]:
    st.write(
        f"共 {len(customer_duplicates):,} 筆"
    )

    st.dataframe(
        customer_duplicates,
        use_container_width=True,
        hide_index=True,
        height=520,
    )


# SKU Duplicate Mapping
with tabs[6]:
    st.write(
        f"共 {len(sku_duplicates):,} 筆"
    )

    st.dataframe(
        sku_duplicates,
        use_container_width=True,
        hide_index=True,
        height=520,
    )


# ============================================================
# Export Data
# ============================================================
st.divider()

# ============================================================
# STEP 5 - EXPORT
# ============================================================

section_header(
    "5",
    "Export (🎃currently under construction🕸️)",
    "Download final report, CSV and unmapped records",
)

unmapped_mask = (
    detail[
        "CUSTOMER_MAPPING_STATUS"
    ].eq("UNMAPPED")
    |
    detail[
        "SKU_MAPPING_STATUS"
    ].eq("UNMAPPED")
)

unmapped = pd.concat(
    [
        detail.loc[
            unmapped_mask,
            audit_columns,
        ].reset_index(
            drop=True
        ),

        final_report.loc[
            unmapped_mask
        ].reset_index(
            drop=True
        ),
    ],
    axis=1,
)

audit_export = pd.concat(
    [
        detail[
            audit_columns
        ].reset_index(
            drop=True
        ),

        final_report.reset_index(
            drop=True
        ),
    ],
    axis=1,
)


# ============================================================
# Full Workbook
# ============================================================
workbook = excel_bytes({
    "Final Report":
        final_report,

    "Upload Log":
        upload_log,

    "Submission Tracker":
        submission,

    "Unmapped":
        unmapped,

    "Audit Detail":
        audit_export,

    "Customer Map Duplicates":
        customer_duplicates,

    "SKU Map Duplicates":
        sku_duplicates,
})


# ============================================================
# Download Buttons
# ============================================================
export_col1, export_col2, export_col3 = (
    st.columns(3)
)

with export_col1:
    st.download_button(
        "下載完整 Excel",
        data=workbook,
        file_name=(
            "POS_Monthly_Report.xlsx"
        ),
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

with export_col2:
    st.download_button(
        "下載 Final CSV",
        data=(
            final_report
            .to_csv(
                index=False
            )
            .encode(
                "utf-8-sig"
            )
        ),
        file_name=(
            "POS_Monthly_Final.csv"
        ),
        mime="text/csv",
        use_container_width=True,
    )

with export_col3:
    st.download_button(
        "下載 Unmapped Excel",
        data=excel_bytes({
            "Unmapped":
                unmapped
        }),
        file_name=(
            "POS_Unmapped.xlsx"
        ),
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )


# ============================================================
# Protection Notes
# ============================================================
st.divider()

st.subheader(
    "🔎 資料保護邏輯"
)

st.markdown(
    """
- 所有 Template 都從同一個上傳區上傳。
- 系統會自動辨識 Template 格式。
- Customer Mapping 採 LEFT JOIN。
- SKU Mapping 採 LEFT JOIN。
- Unmapped 資料不會被刪除。
- Dashboard 固定核對來源資料筆數與輸出筆數。
- 每一筆資料保留來源檔案與來源列資訊。
- Mapping 重複值會另外顯示供檢查。
- 檔名中的業務代號與合約編號會與 Customer Mapping 核對。
"""
)
