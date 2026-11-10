import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st


import streamlit as st

st.write("Streamlit version:", st.__version__)

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
    "SKU Mapping Status",
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

    raw = (
        series
        .astype(str)
        .str.strip()
        .str.replace(",", "", regex=False)
    )


    # ========================================================
    # Exclude barcode / product-code-like values
    #
    # Example:
    # 4713268580974
    # 6921014044399
    #
    # Pure integer with 8 ~ 14 digits
    # will NOT be treated as quantity.
    # ========================================================

    is_barcode_like = (
        raw
        .str.match(
            r"^\d{8,14}$",
            na=False,
        )
    )


    cleaned = (
        raw
        .str.replace(
            r"(瓶|支|罐|箱|盒|包|桶|組|入|個|PCS|PC|BOTTLE|BOTTLES)",
            "",
            regex=True,
            case=False,
        )
        .str.replace(
            r"\s+",
            "",
            regex=True,
        )
        .str.extract(
            r"([-+]?\d*\.?\d+)",
            expand=False,
        )
    )


    result = pd.to_numeric(
        cleaned,
        errors="coerce",
    )


    # Barcode-like values -> NaN
    result.loc[
        is_barcode_like
    ] = pd.NA


    return result

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


    uploaded_file.seek(0)


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
            "產品品名"
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
    # 駱北 / 全億 ONLY
    # 只保留「單位 = 瓶 / 盒」
    # ========================================================
    
    filename_text = norm_text(
        Path(
            uploaded_file.name
        ).stem
    )
    
    
    if (
        "駱北" in filename_text
        or
        "全億" in filename_text
    ):
    
        unit_col = first_matching_column(
            df,
            [
                "單位",
                "Unit",
            ],
        )
    
    
        if unit_col is None:
    
            raise ValueError(
                "Template 7｜駱北 / 全億 "
                "找不到「單位」欄位。"
            )
    
    
        valid = (
            valid
            &
            df[unit_col]
            .map(norm_text)
            .isin(
                [
                    "瓶",
                    "盒",
                    "組",
                ]
            )
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
# TEMPLATE 8
# 唯君系列
#
# ROW-LEVEL UNIT / QTY DETECTION
#
# IMPORTANT:
#
# - NO BARCODE
# - Unit column can move between rows
# - QTY column can move between rows
# - Blank columns are skipped automatically
#
# Structure:
#
# [Product Description Area]
#           ↓
#        [UNIT]
#           ↓
#     Blank / Blank
#           ↓
#         [QTY]
#           ↓
#     Blank / Amount
#
# Example:
#
# Row A:
# SKU ... 瓶 | Blank | 19 | Blank | 13110
#
# Row B:
# SKU ... Blank | 瓶 | 1 | Blank | 350
#
# Both rows are supported.
# ============================================================

def transform_template_8(uploaded_file):

    # ========================================================
    # 1. Filename Metadata
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


    if (
        sheets is None
        or
        not isinstance(
            sheets,
            dict,
        )
        or
        not sheets
    ):

        raise ValueError(
            "Template 8 找不到任何工作表。"
        )


    # ========================================================
    # 3. Settings
    # ========================================================

    UNIT_VALUES = {
        "瓶",
        "罐",
        "支",
        "箱",
        "盒",
        "組",
    }


    # --------------------------------------------------------
    # Common non-product labels
    # --------------------------------------------------------

    INVALID_SKU_KEYS = {

        norm_key("商品名稱"),
        norm_key("商品"),
        norm_key("品名"),
        norm_key("品項名稱"),
        norm_key("產品名稱"),
        norm_key("產品品名"),
        norm_key("貨品名稱"),

        norm_key("條碼"),
        norm_key("商品條碼"),
        norm_key("Barcode"),

        norm_key("單位"),
        norm_key("數量"),
        norm_key("銷量"),
        norm_key("金額"),
        norm_key("售價"),
        norm_key("小計"),

        norm_key("合計"),
        norm_key("總計"),
        norm_key("TOTAL"),
        norm_key("以下空白"),
    }


    frames = []


    # ========================================================
    # 4. Helpers
    # ========================================================

    def to_number(value):

        """
        Convert ONE cell to numeric.

        Unlike global numeric(),
        this is intentionally simple and row-level.
        """

        if pd.isna(value):

            return pd.NA


        text = norm_text(
            value
        )


        if not text:

            return pd.NA


        text = (
            text
            .replace(",", "")
            .replace("$", "")
            .replace("NT$", "")
            .replace("NTD", "")
        )


        # Remove quantity units if embedded
        text = re.sub(
            r"(瓶|罐|支|箱|盒|組|包|桶|入|個|PCS|PC)",
            "",
            text,
            flags=re.IGNORECASE,
        )


        text = text.strip()


        # Strict numeric first
        number = pd.to_numeric(
            text,
            errors="coerce",
        )


        if pd.notna(
            number
        ):

            return number


        # Fallback:
        # 26瓶 -> 26
        # 2,408元 -> 2408
        match = re.search(
            r"[-+]?\d*\.?\d+",
            text,
        )


        if match:

            return pd.to_numeric(
                match.group(),
                errors="coerce",
            )


        return pd.NA


    # ========================================================

    def is_numeric_only(value):

        """
        Check whether cell is essentially numeric.
        Used to prevent numbers from entering RAW_SKU.
        """

        value_num = to_number(
            value
        )


        if pd.isna(
            value_num
        ):

            return False


        text = norm_text(
            value
        )


        cleaned = (
            text
            .replace(",", "")
            .replace("$", "")
            .strip()
        )


        return bool(
            re.fullmatch(
                r"[-+]?\d*\.?\d+",
                cleaned,
            )
        )


    # ========================================================

    def find_unit_in_row(row):

        """
        Find actual unit position in THIS row.

        Unit position is NOT assumed to be fixed.
        """

        for col_idx in range(
            len(row)
        ):

            value = norm_text(
                row.iloc[
                    col_idx
                ]
            )


            if value in UNIT_VALUES:

                return (
                    col_idx,
                    value,
                )


        return (
            None,
            "",
        )


    # ========================================================

    def find_qty_after_unit(
        row,
        unit_col,
    ):

        """
        Starting immediately after UNIT:

        - skip Blank
        - skip text
        - first valid numeric cell = QTY
        """

        for col_idx in range(
            unit_col + 1,
            len(row),
        ):

            raw_value = (
                row.iloc[
                    col_idx
                ]
            )


            # -----------------------------------------------
            # Blank
            # -----------------------------------------------

            if pd.isna(
                raw_value
            ):

                continue


            text = norm_text(
                raw_value
            )


            if not text:

                continue


            # -----------------------------------------------
            # Numeric
            # -----------------------------------------------

            qty_value = to_number(
                raw_value
            )


            if pd.isna(
                qty_value
            ):

                continue


            # -----------------------------------------------
            # Extra QTY sanity check
            #
            # Avoid obvious giant numbers accidentally
            # becoming QTY.
            #
            # This is intentionally generous.
            # -----------------------------------------------

            if abs(
                float(qty_value)
            ) > 100000:

                continue


            return (
                col_idx,
                qty_value,
            )


        return (
            None,
            pd.NA,
        )


    # ========================================================

    def find_amount_after_qty(
        row,
        qty_col,
    ):

        """
        Amount is audit only.

        Search all numeric cells to the right of QTY.
        Use the LAST valid numeric value.
        """

        amount_col = None
        amount_value = pd.NA


        for col_idx in range(
            qty_col + 1,
            len(row),
        ):

            raw_value = (
                row.iloc[
                    col_idx
                ]
            )


            if pd.isna(
                raw_value
            ):

                continue


            text = norm_text(
                raw_value
            )


            if not text:

                continue


            number = to_number(
                raw_value
            )


            if pd.isna(
                number
            ):

                continue


            amount_col = col_idx
            amount_value = number


        return (
            amount_col,
            amount_value,
        )


    # ========================================================

    def build_sku_before_unit(
        row,
        unit_col,
    ):

        """
        Build SKU from text cells BEFORE UNIT.

        Important:
        - remove blanks
        - remove pure numbers
        - remove known header labels
        - remove units
        - remove adjacent duplicates

        Because 唯君 SKU may span multiple cells.
        """

        parts = []


        for col_idx in range(
            0,
            unit_col,
        ):

            raw_value = (
                row.iloc[
                    col_idx
                ]
            )


            value = norm_text(
                raw_value
            )


            if not value:

                continue


            # -----------------------------------------------
            # Unit
            # -----------------------------------------------

            if value in UNIT_VALUES:

                continue


            # -----------------------------------------------
            # Pure numbers
            #
            # e.g.
            # 700
            # 40
            # 0.4
            # -----------------------------------------------

            if is_numeric_only(
                value
            ):

                continue


            # -----------------------------------------------
            # Header / Summary Labels
            # -----------------------------------------------

            value_key = norm_key(
                value
            )


            if value_key in INVALID_SKU_KEYS:

                continue


            # -----------------------------------------------
            # Skip obvious date-like values
            # -----------------------------------------------

            if re.fullmatch(
                r"\d{4}[-/]\d{1,2}[-/]\d{1,2}",
                value,
            ):

                continue


            parts.append(
                value
            )


        # ====================================================
        # Remove adjacent duplicates
        # ====================================================

        cleaned_parts = []


        for part in parts:

            if (
                not cleaned_parts
                or
                part != cleaned_parts[-1]
            ):

                cleaned_parts.append(
                    part
                )


        sku = norm_text(
            " ".join(
                cleaned_parts
            )
        )


        return sku


    # ========================================================
    # 5. Process Every Sheet
    # ========================================================

    for (
        sheet_name,
        raw
    ) in sheets.items():

        # ----------------------------------------------------
        # Validate Sheet
        # ----------------------------------------------------

        if (
            raw is None
            or
            raw.empty
        ):

            continue


        raw = raw.copy()


        records = []


        # ====================================================
        # 6. Process EVERY ROW independently
        # ====================================================

        for row_idx in range(
            len(raw)
        ):

            row = (
                raw.iloc[
                    row_idx
                ]
            )


            # =================================================
            # STEP A
            # Find Unit in THIS row
            # =================================================

            (
                unit_col,
                unit_value,
            ) = find_unit_in_row(
                row
            )


            if unit_col is None:

                continue


            # =================================================
            # STEP B
            # Find QTY AFTER Unit
            # =================================================

            (
                qty_col,
                qty_value,
            ) = find_qty_after_unit(
                row,
                unit_col,
            )


            if (
                qty_col is None
                or
                pd.isna(
                    qty_value
                )
            ):

                continue


            # =================================================
            # STEP C
            # Build SKU BEFORE Unit
            # =================================================

            sku = build_sku_before_unit(
                row,
                unit_col,
            )


            if not sku:

                continue


            # -------------------------------------------------
            # Basic SKU Validation
            # -------------------------------------------------

            sku_key = norm_key(
                sku
            )


            if (
                not sku_key
                or
                sku_key in INVALID_SKU_KEYS
            ):

                continue


            # =================================================
            # STEP D
            # Find Amount after QTY
            # =================================================

            (
                amount_col,
                amount_value,
            ) = find_amount_after_qty(
                row,
                qty_col,
            )


            # =================================================
            # STEP E
            # Save Record
            # =================================================

            records.append({

                # ---------------------------------------------
                # Product
                # ---------------------------------------------

                "RAW_PRODUCT_CODE":
                    "",

                "RAW_SKU":
                    sku,


                # ---------------------------------------------
                # Quantity
                # ---------------------------------------------

                "QTY":
                    qty_value,


                # ---------------------------------------------
                # Raw Price
                # ---------------------------------------------

                "RAW_TOTAL_PRICE":
                    amount_value,


                # ---------------------------------------------
                # Audit
                # ---------------------------------------------

                "SOURCE_ROW":
                    row_idx + 1,

                "SOURCE_SHEET":
                    sheet_name,

                "SOURCE_UNIT":
                    unit_value,

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

            sheet_df = pd.DataFrame(
                records
            )


            frames.append(
                sheet_df
            )


    # ========================================================
    # 8. Validation
    # ========================================================

    if not frames:

        raise ValueError(
            "Template 8 所有 Sheet 都沒有找到有效的 "
            "SKU / Unit / QTY 資料。"
        )


    # ========================================================
    # 9. Combine All Sheets
    # ========================================================

    data = pd.concat(
        frames,
        ignore_index=True,
        sort=False,
    )


    # ========================================================
    # 10. Clean Data
    # ========================================================

    data[
        "RAW_SKU"
    ] = (
        data[
            "RAW_SKU"
        ]
        .map(
            norm_text
        )
    )


    data[
        "QTY"
    ] = pd.to_numeric(
        data[
            "QTY"
        ],
        errors="coerce",
    )


    data[
        "RAW_TOTAL_PRICE"
    ] = pd.to_numeric(
        data[
            "RAW_TOTAL_PRICE"
        ],
        errors="coerce",
    )


    # ========================================================
    # 11. Final Row Validation
    # ========================================================

    valid = (

        data[
            "RAW_SKU"
        ]
        .ne("")

        &

        data[
            "QTY"
        ]
        .notna()

    )


    data = (
        data.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    if data.empty:

        raise ValueError(
            "Template 8 找到結構，"
            "但沒有有效的商品 / QTY 資料。"
        )


    # ========================================================
    # 12. Build Standard Output
    # ========================================================

    output = base_records(
        data,
        meta,
        "T8_WEIJUN_ROW_LEVEL",
        uploaded_file.name,
    )


    # ========================================================
    # 13. Source
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = (
        data[
            "SOURCE_ROW"
        ]
        .values
    )


    output[
        "SOURCE_SHEET"
    ] = (
        data[
            "SOURCE_SHEET"
        ]
        .values
    )


    # ========================================================
    # 14. Customer
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
    # 15. Product
    #
    # NO BARCODE
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = ""


    output[
        "RAW_SKU"
    ] = (
        data[
            "RAW_SKU"
        ]
        .values
    )


    # ========================================================
    # 16. Quantity
    # ========================================================

    output[
        "QTY"
    ] = (
        data[
            "QTY"
        ]
        .values
    )


    # ========================================================
    # 17. Raw Amount
    #
    # Audit only
    # ========================================================

    output[
        "RAW_TOTAL_PRICE"
    ] = (
        data[
            "RAW_TOTAL_PRICE"
        ]
        .values
    )


    output[
        "RAW_UNIT_PRICE"
    ] = pd.NA


    # ========================================================
    # 18. Final Price
    #
    # Final Report uses SKU Mapping RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 19. Optional Raw Fields
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
    # 20. Audit Fields
    # ========================================================

    output[
        "SOURCE_UNIT"
    ] = (
        data[
            "SOURCE_UNIT"
        ]
        .values
    )


    output[
        "SOURCE_UNIT_COLUMN"
    ] = (
        data[
            "SOURCE_UNIT_COLUMN"
        ]
        .values
    )


    output[
        "SOURCE_QTY_COLUMN"
    ] = (
        data[
            "SOURCE_QTY_COLUMN"
        ]
        .values
    )


    output[
        "SOURCE_AMOUNT_COLUMN"
    ] = (
        data[
            "SOURCE_AMOUNT_COLUMN"
        ]
        .values
    )


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    # ========================================================
    # 21. Clean attrs
    # ========================================================

    output.attrs = {}


    # ========================================================
    # 22. Final Validation
    # ========================================================

    if output.empty:

        raise ValueError(
            "Template 8 最終沒有有效資料。"
        )


    # ========================================================
    # 23. Return
    # ========================================================

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
# 橡木桶
#
# Logic:
# - Product Code = Column A
# - SKU          = Column B
# - QTY          = Current month column
#
# Example:
# filename period = 202609
# → target month = SEP
# → QTY = SEP column
# ============================================================

def transform_template_13(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )


    # ========================================================
    # 2. Get Target Month From Filename
    # ========================================================

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


    target_month = pd.to_numeric(
        meta.get(
            "MONTH"
        ),
        errors="coerce",
    )


    if pd.isna(
        target_month
    ):

        raise ValueError(
            "Template 13｜橡木桶 "
            "無法從檔名取得月份。"
        )


    target_month = int(
        target_month
    )


    if (
        target_month
        not in month_map
    ):

        raise ValueError(
            "Template 13｜橡木桶 "
            f"月份不正確：{target_month}"
        )


    target_month_col = (
        month_map[
            target_month
        ]
    )


    # ========================================================
    # 3. Read Excel Raw
    # ========================================================

    uploaded_file.seek(0)

    raw = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        header=None,
        engine="calamine",
    )


    if raw.empty:

        raise ValueError(
            "Template 13｜橡木桶 "
            "檔案沒有資料。"
        )


    # ========================================================
    # 4. Find Header Row
    #
    # Look for:
    # JAN / FEB / ... / DEC
    #
    # Example actual row:
    #
    # NO. | in | | JAN | FEB | ... | SEP | ...
    # ========================================================

    header_row = None


    month_headers = {
        "JAN",
        "FEB",
        "MAR",
        "APR",
        "MAY",
        "JUN",
        "JUL",
        "AUG",
        "SEP",
        "OCT",
        "NOV",
        "DEC",
    }


    for i in range(
        min(
            30,
            len(raw),
        )
    ):

        values = [
            norm_text(value)
            .upper()

            for value
            in raw.iloc[
                i
            ].tolist()
        ]


        found_months = (
            month_headers
            .intersection(
                set(values)
            )
        )


        # 至少找到 6 個月份欄位
        # 才認為這列是 month header
        if (
            len(
                found_months
            )
            >= 6
        ):

            header_row = i

            break


    if (
        header_row
        is None
    ):

        raise ValueError(
            "Template 13｜橡木桶 "
            "找不到 JAN～DEC 月份 Header。"
        )


    # ========================================================
    # 5. Get Header Values
    # ========================================================

    headers = [
        norm_text(value)
        .upper()

        for value
        in raw.iloc[
            header_row
        ].tolist()
    ]


    # ========================================================
    # 6. Find Target Month Column Index
    # ========================================================

    try:

        qty_col_idx = (
            headers.index(
                target_month_col
            )
        )

    except ValueError:

        raise ValueError(
            "Template 13｜橡木桶 "
            f"找不到當月份欄位："
            f"{target_month_col}"
        )


    # ========================================================
    # 7. Data Starts After Header
    # ========================================================

    df = (
        raw.iloc[
            header_row + 1:
        ]
        .copy()
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 8. Fixed Product Columns
    #
    # Column A = Product Code
    # Column B = SKU
    # ========================================================

    raw_product_code = (
        df.iloc[
            :,
            0
        ]
        .map(
            norm_text
        )
    )


    raw_sku = (
        df.iloc[
            :,
            1
        ]
        .map(
            norm_text
        )
    )


    qty = numeric(
        df.iloc[
            :,
            qty_col_idx
        ]
    )


    # ========================================================
    # 9. Valid Product Rows
    #
    # Keep:
    # - Product Code exists
    # - SKU exists
    # - QTY is numeric
    #
    # Remove:
    # - Bootles
    # - TOTAL
    # - blank rows
    # ========================================================

    valid = (
        raw_product_code
        .ne("")
        &
        raw_sku
        .ne("")
        &
        qty
        .notna()
    )


    # Remove non-product rows
    invalid_keywords = [
        "BOOTLES",
        "BOTTLES",
        "TOTAL",
        "NO.",
    ]


    product_code_upper = (
        raw_product_code
        .str
        .upper()
    )


    for keyword in (
        invalid_keywords
    ):

        valid = (
            valid
            &
            ~product_code_upper
            .str
            .startswith(
                keyword
            )
        )


    # ========================================================
    # 10. Apply Filter
    # ========================================================

    df = (
        df.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    raw_product_code = (
        raw_product_code.loc[
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
    # 11. Build Output
    # ========================================================

    output = base_records(
        df,
        meta,
        "T13_OAK_BARREL",
        uploaded_file.name,
    )


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


    output[
        "RAW_PRODUCT_CODE"
    ] = (
        raw_product_code
    )


    output[
        "RAW_SKU"
    ] = (
        raw_sku
    )


    output[
        "QTY"
    ] = (
        qty
    )


    # ========================================================
    # 12. Price
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
    # 13. Audit Information
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = (
        target_month_col
    )


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
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
# ============================================================
# TEMPLATE 16
# Store Sales Matrix
#
# Rule:
#
# Product Code = 產品編號 / 商品編號 / 品號...
# SKU          = 產品品名 / 商品名稱 / 品名...
# QTY          = SKU column immediately to the RIGHT
#
# Example:
#
# 產品編號 | 產品品名             | 東門店
# F120     | 亞伯樂 雪莉原桶-76桶 | 3
#
# RAW_PRODUCT_CODE = F120
# RAW_SKU          = 亞伯樂 雪莉原桶-76桶
# QTY              = 3
#
# Price:
# Final Report uses SKU Mapping RSP
# ============================================================

def transform_template_16(uploaded_file):

    # ========================================================
    # 1. Parse Filename
    # ========================================================

    meta = parse_filename(
        uploaded_file.name
    )

    suffix = Path(
        uploaded_file.name
    ).suffix.lower()


    # ========================================================
    # 2. Read Raw File
    # ========================================================

    uploaded_file.seek(0)


    if suffix == ".csv":

        raw = read_csv_flexible(
            uploaded_file,
            header=None,
        )


    elif suffix == ".xlsx":

        uploaded_file.seek(0)

        raw = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            header=None,
            engine="openpyxl",
        )


    elif suffix == ".xls":

        try:

            uploaded_file.seek(0)

            raw = pd.read_excel(
                uploaded_file,
                sheet_name=0,
                header=None,
                engine="xlrd",
            )

        except Exception:

            uploaded_file.seek(0)

            raw = pd.read_excel(
                uploaded_file,
                sheet_name=0,
                header=None,
            )


    else:

        raise ValueError(
            f"Template 16 不支援此檔案格式：{suffix}"
        )


    if raw is None:

        raise ValueError(
            "Template 16 讀取結果為 None。"
        )


    if raw.empty:

        raise ValueError(
            "Template 16 檔案沒有資料。"
        )


    # ========================================================
    # 3. Header Aliases
    # ========================================================

    product_code_aliases = [
        "產品編號",
        "產品代號",
        "商品編號",
        "商品代號",
        "商品號碼",
        "貨品編號",
        "貨品代號",
        "貨品號碼",
        "品號",
        "SKU Code",
        "SKU code",
    ]


    product_name_aliases = [
        "產品品名",
        "產品名稱",
        "商品名稱",
        "商品品名",
        "貨品名稱",
        "品項名稱",
        "品名",
        "SKU Name",
        "Product Name",
    ]


    # ========================================================
    # 4. Find Header Row
    #
    # Only need:
    #
    # Product Code
    # +
    # Product Name
    #
    # QTY header does NOT matter.
    # ========================================================

    header_row = None


    for i in range(
        min(
            30,
            len(raw),
        )
    ):

        row_values = [
            norm_text(value)
            for value
            in raw.iloc[i].tolist()
        ]


        normalized_values = [
            norm_key(value)
            for value
            in row_values
            if norm_text(value)
        ]


        has_product_code = any(
            norm_key(alias)
            in normalized_values

            for alias
            in product_code_aliases
        )


        has_product_name = any(
            norm_key(alias)
            in normalized_values

            for alias
            in product_name_aliases
        )


        if (
            has_product_code
            and
            has_product_name
        ):

            header_row = i

            break


    if header_row is None:

        raise ValueError(
            "Template 16 找不到「產品編號 + 產品品名」類型 Header。"
        )


    # ========================================================
    # 5. Read Again Using Correct Header
    # ========================================================

    uploaded_file.seek(0)


    if suffix == ".csv":

        df = read_csv_flexible(
            uploaded_file,
            header=header_row,
        )


    elif suffix == ".xlsx":

        uploaded_file.seek(0)

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            header=header_row,
            engine="openpyxl",
        )


    elif suffix == ".xls":

        try:

            uploaded_file.seek(0)

            df = pd.read_excel(
                uploaded_file,
                sheet_name=0,
                header=header_row,
                engine="xlrd",
            )

        except Exception:

            uploaded_file.seek(0)

            df = pd.read_excel(
                uploaded_file,
                sheet_name=0,
                header=header_row,
            )


    else:

        raise ValueError(
            f"Template 16 不支援此檔案格式：{suffix}"
        )


    if df is None:

        raise ValueError(
            "Template 16 Header 重新讀取結果為 None。"
        )


    # ========================================================
    # 6. Clean Column Names
    # ========================================================

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


    if df.empty:

        raise ValueError(
            "Template 16 Header 後沒有資料。"
        )


    # ========================================================
    # 7. Find Product Code Column
    # ========================================================

    product_code_col = first_matching_column(
        df,
        product_code_aliases,
    )


    if product_code_col is None:

        raise ValueError(
            "Template 16 找不到產品編號欄位。"
            f" Available Columns={list(df.columns)}"
        )


    # ========================================================
    # 8. Find SKU Column
    # ========================================================

    sku_col = first_matching_column(
        df,
        product_name_aliases,
    )


    if sku_col is None:

        raise ValueError(
            "Template 16 找不到產品品名欄位。"
            f" Available Columns={list(df.columns)}"
        )


    # ========================================================
    # 9. QTY Column
    #
    # IMPORTANT:
    #
    # Do NOT use store name.
    #
    # QTY = column immediately RIGHT of SKU column.
    #
    # Example:
    #
    # 產品編號 | 產品品名 | 東門店
    #
    # SKU index = 1
    # QTY index = 2
    # ========================================================

    columns = list(
        df.columns
    )


    try:

        sku_col_index = columns.index(
            sku_col
        )


    except ValueError:

        raise ValueError(
            "Template 16 無法定位產品品名欄位位置。"
        )


    qty_col_index = (
        sku_col_index
        + 1
    )


    if (
        qty_col_index
        >= len(columns)
    ):

        raise ValueError(
            "Template 16 產品品名右邊沒有欄位，"
            "因此無法取得銷量。"
        )


    qty_col = columns[
        qty_col_index
    ]


    # ========================================================
    # 10. Prepare Raw Product Code
    # ========================================================

    raw_product_code = (
        df[
            product_code_col
        ]
        .map(
            norm_text
        )
    )


    # ========================================================
    # 11. Prepare Raw SKU
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
    # 12. Prepare Quantity
    # ========================================================

    qty = numeric(
        df[
            qty_col
        ]
    )


    # ========================================================
    # 13. Valid Rows
    #
    # Keep:
    #
    # SKU not blank
    # QTY numeric
    # ========================================================

    valid = (
        raw_sku
        .ne("")
        &
        qty
        .notna()
    )


    # ========================================================
    # 14. Remove Obvious Header / Total Rows
    # ========================================================

    product_code_key = (
        raw_product_code
        .map(
            norm_key
        )
    )


    invalid_code_keys = {
        norm_key("TOTAL"),
        norm_key("合計"),
        norm_key("總計"),
        norm_key("NO"),
        norm_key("NO."),
        norm_key("產品編號"),
        norm_key("商品編號"),
        norm_key("品號"),
    }


    valid = (
        valid
        &
        ~product_code_key.isin(
            invalid_code_keys
        )
    )


    # ========================================================
    # 15. Apply Filter
    # ========================================================

    df = (
        df.loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    raw_product_code = (
        raw_product_code.loc[
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
    # 16. Validate Final Data
    # ========================================================

    if df.empty:

        raise ValueError(
            "Template 16 已找到產品品名與右側 QTY 欄位，"
            "但沒有可使用的 SKU / QTY 資料。"
            f" SKU Column={sku_col}"
            f"｜QTY Column={qty_col}"
        )


    # ========================================================
    # 17. Build Standard Output
    # ========================================================

    output = base_records(
        df,
        meta,
        "T16_STORE_SALES_MATRIX",
        uploaded_file.name,
    )


    # ========================================================
    # 18. Source Row
    # ========================================================

    output[
        "SOURCE_ROW"
    ] = range(
        header_row + 2,
        header_row + 2 + len(df),
    )


    # ========================================================
    # 19. Customer
    #
    # Still use filename metadata for customer mapping.
    #
    # But it is NOT used for finding QTY anymore.
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
    # 20. Product Code
    # ========================================================

    output[
        "RAW_PRODUCT_CODE"
    ] = raw_product_code


    # ========================================================
    # 21. SKU
    # ========================================================

    output[
        "RAW_SKU"
    ] = raw_sku


    # ========================================================
    # 22. Quantity
    # ========================================================

    output[
        "QTY"
    ] = qty


    # ========================================================
    # 23. Raw Source Price
    #
    # Not used in Template 16
    # ========================================================

    output[
        "RAW_UNIT_PRICE"
    ] = pd.NA


    output[
        "RAW_TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 24. Final Price
    #
    # Final report uses SKU Mapping RSP
    # ========================================================

    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    # ========================================================
    # 25. Optional Raw Fields
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
    # 26. Audit
    # ========================================================

    output[
        "SOURCE_QTY_COLUMN"
    ] = str(
        qty_col
    )


    output[
        "SOURCE_PRICE_COLUMN"
    ] = (
        "RSP_FROM_SKU_MAPPING"
    )


    # ========================================================
    # 27. Clean attrs
    # ========================================================

    output.attrs = {}


    # ========================================================
    # 28. Return
    # ========================================================

    return output

# ============================================================
# AUTO DETECT TEMPLATE TYPE
# ============================================================

def detect_template_type(uploaded_file):

    # ========================================================
    # 0. Basic File Information
    # ========================================================

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
    # 美滿商行
    #
    # Old .xls special format.
    # Must detect before any Excel reader.
    # ========================================================

    if (
        suffix == ".xls"
        and
        "美滿商行" in filename_text
    ):

        return "TEMPLATE_14"


    # ========================================================
    # TEMPLATE 13
    # 橡木桶
    #
    # Filename based.
    # ========================================================

    if (
        "橡木桶"
        in filename_text
    ):

        return "TEMPLATE_13"


    # ========================================================
    # TEMPLATE 16
    # 合歡
    #
    # Confirmed Store Sales Matrix.
    #
    # IMPORTANT:
    # Detect by filename BEFORE Excel parsing.
    # ========================================================

    if (
        "合歡"
        in filename_text
    ):

        return "TEMPLATE_16"


    # ========================================================
    # CSV FILES
    # ========================================================

    if suffix == ".csv":

        # ====================================================
        # CSV Template 1
        #
        # Standard POS
        # Header = first row
        # ====================================================

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
                    "品名規格",
                ],
            )


            qty_col = first_matching_column(
                df,
                [
                    "銷量 Quantity",
                    "Quantity",
                    "銷量",
                    "銷售量",
                    "數量",
                    "瓶數",
                    "淨銷數量",
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


        # ====================================================
        # CSV Template 7
        #
        # Generic SKU + QTY
        # Header may appear within first 30 rows
        # ====================================================

        try:

            uploaded_file.seek(0)

            raw_csv = read_csv_flexible(
                uploaded_file,
                header=None,
            )


            template_7_header_row = (
                find_header_row(
                    raw_csv,
                    required_fields=[
                        "SKU",
                        "QTY",
                    ],
                    max_rows=30,
                )
            )


            if (
                template_7_header_row
                is not None
            ):

                return "TEMPLATE_7"


        except Exception:

            pass


        return "UNKNOWN"


    # ========================================================
    # ONLY EXCEL BELOW
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
    # Read First Sheet as Raw
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
    # 商品貨號 + 商品名稱
    # ========================================================

    try:

        first_rows_text = " ".join(
            norm_text(value)
            for value
            in raw.iloc[:20]
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
                in raw.iloc[i].tolist()
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
    # POS 公版 / 康尼克
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
    # 松賀
    #
    # Fixed 8-column no-header structure
    # ========================================================

    try:

        if (
            "松賀"
            in filename_text
            and
            raw.shape[1] >= 8
        ):

            return "TEMPLATE_6"


    except Exception:

        pass


    # ========================================================
    # TEMPLATE 8
    # 唯君
    # ========================================================

    if (
        "唯君"
        in filename_text
    ):

        return "TEMPLATE_8"


    # ========================================================
    # TEMPLATE 3A
    # 國泰 Multi-sheet
    #
    # 品名規格 + 數量 / 數量2
    # ========================================================

    if (
        len(
            sheet_names
        )
        > 1
    ):

        try:

            uploaded_file.seek(0)

            first_sheet = pd.read_excel(
                uploaded_file,
                sheet_name=sheet_names[0],
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
    # 商品代號 + 商品名稱 + 實銷
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
                in raw.iloc[i].tolist()
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
    # 貨品名稱 + 數量
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
                in raw.iloc[i].tolist()
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
    # 新店李 / Diageo
    #
    # 品號 + 品名規格 + 淨銷數量
    # ========================================================

    try:

        template_9_header_found = False


        for i in range(
            min(
                20,
                len(raw),
            )
        ):

            normalized_values = [
                norm_key(value)
                for value
                in raw.iloc[i].tolist()
                if norm_text(value)
            ]


            has_product_code = (
                norm_key(
                    "品號"
                )
                in normalized_values
            )


            has_sku = (
                norm_key(
                    "品名規格"
                )
                in normalized_values
            )


            has_qty = (
                norm_key(
                    "淨銷數量"
                )
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
    # 名酒集
    #
    # Column A = SKU
    # Column B = QTY
    # ========================================================

    try:

        if (
            "名酒集"
            in filename_text
            and
            raw.shape[1] >= 2
        ):

            return "TEMPLATE_10"


    except Exception:

        pass


    # ========================================================
    # TEMPLATE 11
    # 酒江村
    #
    # 國際條碼 + 品牌 + 數量
    # ========================================================

    try:

        normalized_columns = [
            norm_key(value)
            for value
            in raw.iloc[0].tolist()
            if norm_text(value)
        ]


        if (
            norm_key(
                "國際條碼"
            )
            in normalized_columns
            and
            norm_key(
                "品牌"
            )
            in normalized_columns
            and
            norm_key(
                "數量"
            )
            in normalized_columns
        ):

            return "TEMPLATE_11"


    except Exception:

        pass


    # ========================================================
    # TEMPLATE 12
    # 廣茂隆
    # ========================================================

    try:

        first_row = [
            norm_key(value)
            for value
            in raw.iloc[0].tolist()
            if norm_text(value)
        ]


        if (
            "廣茂隆"
            in filename_text
            and
            any(
                norm_key(
                    "品項名稱"
                )
                in value

                for value
                in first_row
            )
            and
            raw.shape[1] >= 10
        ):

            return "TEMPLATE_12"


    except Exception:

        pass


    # ========================================================
    # TEMPLATE 15
    # 精湛酒藏 Multi-sheet
    # ========================================================

    try:

        if (
            "精湛酒藏"
            in filename_text
            and
            len(
                sheet_names
            )
            > 1
        ):

            return "TEMPLATE_15"


    except Exception:

        pass


    # ========================================================
    # TEMPLATE 16
    # Store Sales Matrix
    #
    # Example:
    #
    # 產品編號 | 產品品名 | 合歡 | ...
    #
    # No normal QTY header.
    #
    # IMPORTANT:
    # T16 must be checked before T7,
    # but only when the store header can be matched.
    # ========================================================

    try:

        product_code_aliases = [
            "產品編號",
            "商品編號",
            "商品代號",
            "商品號碼",
            "貨品編號",
            "貨品代號",
            "貨品號碼",
            "品號",
        ]


        product_name_aliases = [
            "產品品名",
            "產品名稱",
            "商品名稱",
            "品名",
            "品項名稱",
        ]


        template_16_header_row = None


        for i in range(
            min(
                30,
                len(raw),
            )
        ):

            header_values_raw = [
                norm_text(value)
                for value
                in raw.iloc[i].tolist()
            ]


            header_values_norm = [
                norm_key(value)
                for value
                in header_values_raw
                if norm_text(value)
            ]


            has_product_code = any(
                norm_key(alias)
                in header_values_norm

                for alias
                in product_code_aliases
            )


            has_product_name = any(
                norm_key(alias)
                in header_values_norm

                for alias
                in product_name_aliases
            )


            if (
                has_product_code
                and
                has_product_name
            ):

                template_16_header_row = i

                break


        if (
            template_16_header_row
            is not None
        ):

            header_values_raw = [
                norm_text(value)
                for value
                in raw.iloc[
                    template_16_header_row
                ].tolist()
            ]


            header_values_norm = [
                norm_key(value)
                for value
                in header_values_raw
                if norm_text(value)
            ]


            # ------------------------------------------------
            # If normal QTY header exists,
            # this is probably Template 7 instead.
            # ------------------------------------------------

            has_normal_qty_header = any(
                norm_key(alias)
                in header_values_norm

                for alias
                in STANDARD_HEADER_MAPPING[
                    "QTY"
                ]
            )


            if not has_normal_qty_header:

                meta = parse_filename(
                    uploaded_file.name
                )


                target_store = norm_text(
                    meta[
                        "FILE_RAWDATA_NAME"
                    ]
                )


                target_store_key = norm_key(
                    target_store
                )


                # --------------------------------------------
                # Store column must match filename store
                # --------------------------------------------

                store_header_found = any(
                    target_store_key
                    and
                    (
                        target_store_key
                        ==
                        norm_key(value)

                        or

                        target_store_key
                        in
                        norm_key(value)

                        or

                        (
                            norm_key(value)
                            and
                            norm_key(value)
                            in
                            target_store_key
                        )
                    )

                    for value
                    in header_values_raw

                    if norm_text(value)
                )


                if store_header_found:

                    return "TEMPLATE_16"


    except Exception:

        pass


    # ========================================================
    # TEMPLATE 7
    # Generic SKU + QTY
    #
    # Header can appear anywhere in first 30 rows
    #
    # IMPORTANT:
    # Must remain after special templates.
    # ========================================================

    try:

        template_7_header_row = (
            find_header_row(
                raw,
                required_fields=[
                    "SKU",
                    "QTY",
                ],
                max_rows=30,
            )
        )


        if (
            template_7_header_row
            is not None
        ):

            return "TEMPLATE_7"


    except Exception:

        pass


    # ========================================================
    # TEMPLATE 1
    # Standard POS Excel
    #
    # Header normally row 1
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
                "品名規格",
            ],
        )


        qty_col = first_matching_column(
            df,
            [
                "銷量 Quantity",
                "Quantity",
                "銷量",
                "銷售量",
                "數量",
                "瓶數",
                "淨銷數量",
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

def auto_transform_template(
    uploaded_file,
    template_type=None,
):

    uploaded_file.seek(0)

    if template_type is None:

        template_type = detect_template_type(
            uploaded_file
        )


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
            "Template 13｜橡木桶",
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


    elif template_type == "TEMPLATE_16":

        uploaded_file.seek(0)
    
        result = transform_template_16(
            uploaded_file
        )

        return (
            result,
            "Template 16｜Store Sales Matrix",
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
        "銷售量"
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
    # 6. RESPONSIBLE SALES
    #
    # FILE_SALES_ID 不再決定負責業務
    #
    # 負責業務一律來自 Customer Mapping：
    #
    # FILE_OUTLET_NO
    #     ↓
    # Customer Mapping - Outlet No
    #     ↓
    # Sales ID / Sales
    #
    # ========================================================

    merged[
        "SUBMISSION_CUSTOMER_ID"
    ] = (
        merged[
            "FILE_OUTLET_NO"
        ]
        .map(
            norm_id_key
        )
    )


    # --------------------------------------------------------
    # Responsible Sales ID
    # --------------------------------------------------------

    merged[
        "RESPONSIBLE_SALES_ID"
    ] = (
        merged[
            "Sales ID"
        ]
        .fillna("")
        .map(
            norm_id_key
        )
    )


    # --------------------------------------------------------
    # Responsible Sales Name
    # --------------------------------------------------------

    merged[
        "RESPONSIBLE_SALES"
    ] = (
        merged[
            "Sales"
        ]
        .fillna("")
        .map(
            norm_text
        )
    )


    # --------------------------------------------------------
    # Responsible Sales Status
    #
    # Customer Mapping 成功
    # 且有 Sales ID
    # 才算成功判定負責業務
    # --------------------------------------------------------

    merged[
        "RESPONSIBLE_SALES_STATUS"
    ] = merged.apply(
        lambda row:
        (
            "MAPPED"
            if (
                row[
                    "CUSTOMER_MAPPING_STATUS"
                ]
                == "MAPPED"
                and
                norm_id_key(
                    row[
                        "Sales ID"
                    ]
                )
                != ""
            )
            else "UNMAPPED"
        ),
        axis=1,
    )


    # ========================================================
    # 7. FILE SALES ID AUDIT
    #
    # 檔名 Sales ID 只保留 Audit
    # 不再影響：
    #
    # - Customer Mapping
    # - Responsible Sales
    # - Submission Status
    # - 業務繳交進度
    #
    # ========================================================

    merged[
        "FILE_SALES_ID_AUDIT"
    ] = (
        merged[
            "FILE_SALES_ID"
        ]
        .fillna("")
        .map(
            norm_id_key
        )
    )


    # ========================================================
    # 8. CONTRACT CHECK
    #
    # Audit only
    #
    # Contract 不再參與 Submission matching
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
    # 9. OUTLET CHECK
    #
    # 這是最重要的 Customer ID Audit
    #
    # FILE_OUTLET_NO
    # vs
    # Customer Mapping Outlet No
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
    # 10. RAWDATA NAME CHECK
    #
    # Audit only
    #
    # 因為 Customer Mapping 有 fallback by Rawdata Name
    # 所以不一定要完全一致才能視為成功 Mapping
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


    # ========================================================
    # 11. CUSTOMER OWNERSHIP CHECK
    #
    # 最終判定：
    # 這份資料是否成功找到真正負責業務
    # ========================================================

    merged[
        "CUSTOMER_OWNER_STATUS"
    ] = merged.apply(
        lambda row:
        (
            "MAPPED"
            if (
                row[
                    "CUSTOMER_MAPPING_STATUS"
                ]
                == "MAPPED"
                and
                row[
                    "RESPONSIBLE_SALES_STATUS"
                ]
                == "MAPPED"
            )
            else "UNMAPPED"
        ),
        axis=1,
    )


    # ========================================================
    # 12. Return
    # ========================================================

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

    report["SKU Mapping Status"] = (
    detail[
        "SKU_MAPPING_STATUS"
    ]
    .fillna("UNMAPPED")
    .values
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
    # 品項 = 原始 Template 商品名稱
    #
    # IMPORTANT:
    # 現在使用「品項」作為合併依據
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
    # 8. PREPARE QTY
    # ========================================================

    report["銷量(瓶)"] = pd.to_numeric(
        report["銷量(瓶)"],
        errors="coerce",
    )


    # ========================================================
    # 9. MERGE SAME 品項
    #
    # Merge only when these fields are the same:
    #
    # 地區
    # 業務員
    # 合約等級
    # 合約編號
    # 合約名稱
    # 店家編號
    # 店家名稱
    # 品項             <-- CHANGED
    # 年
    # 月
    #
    # 「統一品項名稱」不再是 Group Key
    #
    # Example:
    #
    # 品項=A酒700ml | Mapping=A酒 | QTY=2
    # 品項=A酒700ml | Mapping=A酒 | QTY=3
    #
    # ↓
    #
    # 品項=A酒700ml | Mapping=A酒 | QTY=5
    #
    #
    # But:
    #
    # 品項=A酒700ml | Mapping=A酒
    # 品項=A酒禮盒  | Mapping=A酒
    #
    # 不會合併
    # ========================================================

    group_cols = [
        "地區",
        "業務員",
        "合約等級",
        "合約編號",
        "合約名稱",
        "店家編號",
        "店家名稱",
        "品項",
        "年",
        "月",
    ]


    # ========================================================
    # 10. ONLY MERGE ROWS WITH 品項
    #
    # 避免所有空白 RAW_SKU 被合併成同一筆
    # ========================================================

    sku_mask = (
        report[
            "品項"
        ]
        .map(
            norm_text
        )
        .ne("")
    )


    sku_report = (
        report.loc[
            sku_mask
        ]
        .copy()
    )


    blank_sku_report = (
        report.loc[
            ~sku_mask
        ]
        .copy()
    )


    # ========================================================
    # 11. AGGREGATE SAME 品項
    # ========================================================

    if not sku_report.empty:

       agg_rules = {

    # =================================================
    # Quantity
    # =================================================

    "銷量(瓶)": "sum",


    # =================================================
    # Price
    # =================================================

    "單價": "first",

    "建議售價": "first",


    # =================================================
    # SKU Mapping Status
    #
    # 同一個 品項 group 裡
    # 只要有任何一筆 UNMAPPED
    # 最終就標記為 UNMAPPED
    # =================================================

    "SKU Mapping Status":
        lambda x:
            (
                "UNMAPPED"
                if (
                    x.fillna("UNMAPPED")
                    .astype(str)
                    .str.upper()
                    .eq("UNMAPPED")
                    .any()
                )
                else "MAPPED"
            ),


    # =================================================
    # SKU Mapping Fields
    # =================================================

    "價格帶": "first",

    "製造商": "first",

    "品牌": "first",

    "統一品項名稱": "first",

    "容量": "first",

    "品類": "first",
}


        # ====================================================
        # Recalculate Total Price
        #
        # 總價 = 合併後 QTY × RSP
        # ====================================================

        sku_report[
            "總價"
        ] = (
            pd.to_numeric(
                sku_report[
                    "銷量(瓶)"
                ],
                errors="coerce",
            )
            *
            pd.to_numeric(
                sku_report[
                    "單價"
                ],
                errors="coerce",
            )
        )


    # ========================================================
    # 12. COMBINE
    #
    # 有品項：
    # 同品項進行 aggregation
    #
    # 空白品項：
    # 保留原始 row，不合併
    # ========================================================

    report = pd.concat(
        [
            sku_report,
            blank_sku_report,
        ],
        ignore_index=True,
        sort=False,
    )


    # ========================================================
    # 13. FORCE FINAL COLUMN ORDER
    # ========================================================

    report = report[
        OUTPUT_COLUMNS
    ].copy()


    # ========================================================
    # 14. OPTIONAL SORT
    # ========================================================

    report = (
        report
        .sort_values(
            by=[
                "年",
                "月",
                "店家編號",
                "品項",
            ],
            na_position="last",
            kind="stable",
        )
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # 15. CLEAN ATTRS
    # ========================================================

    report.attrs = {}


    # ========================================================
    # 16. RETURN
    # ========================================================

    return report
# ============================================================
# Excel Export
# ============================================================
# ============================================================
# Excel Export
# Safe Version
#
# Removes characters that Excel / openpyxl cannot write.
# ============================================================

import io
import re
import pandas as pd


# ============================================================
# Illegal XML / Excel Characters
#
# Excel XML 不允許：
# \x00-\x08
# \x0B-\x0C
# \x0E-\x1F
# ============================================================

ILLEGAL_EXCEL_CHARACTERS_RE = re.compile(
    r"[\x00-\x08\x0B\x0C\x0E-\x1F]"
)


def clean_excel_value(value):

    # --------------------------------------------------------
    # Keep NaN / None
    # --------------------------------------------------------

    if pd.isna(value):

        return value


    # --------------------------------------------------------
    # Only clean strings
    # --------------------------------------------------------

    if isinstance(
        value,
        str,
    ):

        value = (
            ILLEGAL_EXCEL_CHARACTERS_RE
            .sub(
                "",
                value,
            )
        )


        # Excel cell text limit = 32,767 characters
        if len(value) > 32767:

            value = value[:32767]


    return value


# ============================================================
# Clean DataFrame Before Excel Export
# ============================================================

def clean_dataframe_for_excel(df):

    if df is None:

        return pd.DataFrame()


    if not isinstance(
        df,
        pd.DataFrame,
    ):

        return pd.DataFrame(
            df
        )


    cleaned = df.copy()


    # --------------------------------------------------------
    # Clean column names
    # --------------------------------------------------------

    cleaned.columns = [
        clean_excel_value(
            str(column)
        )
        for column
        in cleaned.columns
    ]


    # --------------------------------------------------------
    # Clean object / string columns only
    # --------------------------------------------------------

    for column in cleaned.columns:

        if (
            pd.api.types.is_object_dtype(
                cleaned[column]
            )
            or
            pd.api.types.is_string_dtype(
                cleaned[column]
            )
        ):

            cleaned[column] = (
                cleaned[column]
                .map(
                    clean_excel_value
                )
            )


    cleaned.attrs = {}


    return cleaned


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


            # =================================================
            # Clean Sheet Name
            # =================================================

            safe_sheet_name = (
                clean_excel_value(
                    str(
                        sheet_name
                    )
                )
            )


            # Excel sheet names cannot contain:
            # : \ / ? * [ ]
            safe_sheet_name = re.sub(
                r'[:\\/?*\[\]]',
                "_",
                safe_sheet_name,
            )


            safe_sheet_name = (
                safe_sheet_name[:31]
            )


            if not safe_sheet_name:

                safe_sheet_name = "Sheet"


            # =================================================
            # Clean DataFrame
            # =================================================

            safe_dataframe = (
                clean_dataframe_for_excel(
                    dataframe
                )
            )


            # =================================================
            # Write Excel
            # =================================================

            safe_dataframe.to_excel(
                writer,
                sheet_name=safe_sheet_name,
                index=False,
            )


    buffer.seek(0)


    return buffer.getvalue()


# ============================================================
# ============================================================
import io
import zipfile
from pathlib import Path

# ============================================================
# STEP 1 - Upload Monthly Template
# ZIP + Excel/CSV Dual Mode
# ============================================================

section_header(
    "1",
    "Upload monthly POS templates",
    "上傳當月份資料",
)


upload_mode = st.radio(
    "Upload Method",
    [
        "📦 ZIP File (批次上傳)",
        "📄 Excel / CSV (多檔上傳)",
    ],
    horizontal=True,
)


template_files = []


# ============================================================
# MODE A - ZIP Upload
# ============================================================

if upload_mode == "📦 ZIP File (批次上傳)":

    zip_file = st.file_uploader(
        "上傳當月份 Upload POS ZIP File",
        type=["zip"],
        key="pos_zip",
        on_change="ignore",
    )


    if zip_file is not None:

        try:

            zip_file.seek(0)

            with zipfile.ZipFile(
                zip_file,
                "r",
            ) as z:

                supported_files = []

                ignored_files = []


                for file_info in z.infolist():

                    file_name = file_info.filename


                    # ========================================
                    # Ignore folders
                    # ========================================

                    if file_info.is_dir():

                        continue


                    # ========================================
                    # Ignore macOS hidden files
                    # ========================================

                    if (
                        "__MACOSX" in file_name
                        or
                        Path(file_name).name.startswith(".")
                    ):

                        continue


                    suffix = (
                        Path(file_name)
                        .suffix
                        .lower()
                    )


                    # ========================================
                    # Supported Formats
                    # ========================================

                    if suffix not in [
                        ".xlsx",
                        ".xls",
                        ".csv",
                    ]:

                        ignored_files.append(
                            file_name
                        )

                        continue


                    # ========================================
                    # Read file bytes from ZIP
                    # ========================================

                    file_bytes = z.read(
                        file_info
                    )


                    # ========================================
                    # Convert to BytesIO
                    # ========================================

                    file_obj = io.BytesIO(
                        file_bytes
                    )


                    # ========================================
                    # Preserve original filename
                    #
                    # IMPORTANT:
                    # parse_filename() depends on this
                    # ========================================

                    file_obj.name = (
                        Path(file_name).name
                    )


                    # Optional size attribute
                    file_obj.size = (
                        len(file_bytes)
                    )


                    supported_files.append(
                        file_obj
                    )


                template_files = (
                    supported_files
                )


            # ================================================
            # ZIP Summary
            # ================================================

            if template_files:

                total_size_mb = (
                    sum(
                        file.size
                        for file
                        in template_files
                    )
                    /
                    1024
                    /
                    1024
                )


                st.success(
                    f"ZIP 解壓縮成功："
                    f"{len(template_files)} 個可處理檔案"
                )


                col1, col2, col3 = st.columns(
                    3
                )


                col1.metric(
                    "可處理檔案",
                    len(template_files),
                )


                col2.metric(
                    "總大小",
                    f"{total_size_mb:.2f} MB",
                )


                col3.metric(
                    "忽略檔案",
                    len(ignored_files),
                )


                # ============================================
                # File Preview
                # ============================================

                with st.expander(
                    "📂 查看 ZIP 內檔案",
                    expanded=False,
                ):

                    preview_rows = []

                    for file in template_files:

                        preview_rows.append({

                            "檔案":
                                file.name,

                            "大小 KB":
                                round(
                                    file.size
                                    /
                                    1024,
                                    1,
                                ),

                            "格式":
                                Path(
                                    file.name
                                ).suffix.lower(),

                        })


                    st.dataframe(
                        pd.DataFrame(
                            preview_rows
                        ),
                        width="stretch",
                        hide_index=True,
                    )


                # ============================================
                # Ignored Files
                # ============================================

                if ignored_files:

                    with st.expander(
                        "⚠️ 被忽略的檔案",
                        expanded=False,
                    ):

                        for file_name in ignored_files:

                            st.write(
                                file_name
                            )


            else:

                st.warning(
                    "ZIP 裡沒有找到可處理的 "
                    "Excel / CSV 檔案。"
                )


        except zipfile.BadZipFile:

            st.error(
                "ZIP File error (檔案無法讀取或已損毀)。"
            )

            st.stop()


        except Exception as e:

            st.error(
                f"ZIP 解壓縮失敗：{e}"
            )

            st.stop()


# ============================================================
# MODE B - Original Multi-file Upload
# ============================================================

else:

    uploaded_files = st.file_uploader(
        "可一次上傳多個 Excel / CSV Template",
        type=[
            "xlsx",
            "xls",
            "csv",
        ],
        accept_multiple_files=True,
        key="templates",
        on_change="ignore",
    )


    if uploaded_files:

        template_files = (
            uploaded_files
        )


        total_template_size_mb = (
            sum(
                file.size
                for file
                in template_files
            )
            /
            1024
            /
            1024
        )


        st.caption(
            f"已選擇 "
            f"{len(template_files)} 個 Template"
            f" ｜ "
            f"總大小 "
            f"{total_template_size_mb:.2f} MB"
        )


    else:

        st.caption(
            "系統會自動辨識不同 Template 格式，"
            "不需要人工選擇格式。"
        )


# ============================================================
# STEP 2 - Upload Mapping
# ============================================================

section_header(
    "2",
    "Upload customer and SKU mapping files",
    "上傳 Mapping 資料",
)

map_col1, map_col2 = st.columns(2)


with map_col1:

    customer_file = st.file_uploader(
        "Customer Mapping",
        type=["xlsx", "xls", "csv"],
        key="customer",
        on_change="ignore",
    )


with map_col2:

    sku_files = st.file_uploader(
        "SKU Mapping",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="sku",
        on_change="ignore",
    )

    if sku_files:

        st.caption(
            f"已選擇 {len(sku_files)} 個 SKU Mapping 檔案"
        )


# ============================================================
# STEP 3 - Start Processing
# ============================================================

section_header(
    "3",
    "Process uploaded POS and mapping files",
    "開始處理資料",
)


# ============================================================
# Persistent State
# ============================================================

if "processing_started" not in st.session_state:

    st.session_state[
        "processing_started"
    ] = False


# ============================================================
# Buttons
# ============================================================


st.markdown(
    """
    <style>
    div.stButton > button[kind="primary"] {
        background-color: #4A4A4A !important;
        color: white !important;
        border: 1px solid #4A4A4A !important;
        font-weight: 600 !important;
    }

    div.stButton > button[kind="primary"]:hover {
        background-color: #3A3A3A !important;
        border-color: #3A3A3A !important;
        color: white !important;
    }

    div.stButton > button[kind="primary"]:active {
        background-color: #2F2F2F !important;
        border-color: #2F2F2F !important;
        color: white !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


col_start, col_reset = st.columns(
    [4, 1]
)


with col_start:

    if st.button(
        "▶️ Start Processing data",
        type="primary",
        width="stretch",
    ):

        st.session_state[
            "processing_started"
        ] = True


with col_reset:

    if st.button(
        "🔄 Reset",
        width="stretch",
    ):

        st.session_state[
            "processing_started"
        ] = False

        st.rerun()


# ============================================================
# Stop Before Processing
# ============================================================

if not st.session_state[
    "processing_started"
]:

    st.info(
        "Please make sure that all the mapping files have been uploaded before starting the process."
    )

    st.stop()


# ============================================================
# Validate Files
# ============================================================

if not template_files:

    st.error(
        "請至少上傳一個當月份 Template。"
    )

    st.stop()


if customer_file is None:

    st.error(
        "請上傳 Customer Mapping。"
    )

    st.stop()


if not sku_files:

    st.error(
        "請至少上傳一個 SKU Mapping。"
    )

    st.stop()


# ============================================================
# Load Customer Mapping
# ============================================================

try:

    customer_file.seek(0)

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
# Load Multiple SKU Mapping Files
# ============================================================

try:

    sku_frames = []
    sku_file_logs = []


    for sku_file in sku_files:

        try:

            sku_file.seek(0)

            sku_temp = read_mapping_file(
                sku_file,
                preferred_sheet="SKU data",
            )


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


            sku_temp = sku_temp.copy()


            sku_temp[
                "SKU_MAPPING_SOURCE_FILE"
            ] = sku_file.name


            sku_frames.append(
                sku_temp
            )


            sku_file_logs.append({

                "檔案":
                    sku_file.name,

                "狀態":
                    "SUCCESS",

                "筆數":
                    len(sku_temp),

                "錯誤":
                    "",

            })


        except Exception as file_error:

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
    # Create Log
    # ========================================================

    sku_upload_log = pd.DataFrame(
        sku_file_logs
    )


    # ========================================================
    # At least one file successful
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
    # Combine
    # ========================================================

    sku_raw = pd.concat(
        sku_frames,
        ignore_index=True,
        sort=False,
    )


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

    st.stop()


# ============================================================
# SKU Mapping Upload Status
# ============================================================

if "sku_upload_log" in locals():

    with st.expander(
        "📦 SKU Mapping 上傳狀況",
        expanded=False,
    ):

        st.dataframe(
            sku_upload_log,
            width="stretch",
            hide_index=True,
        )


# ============================================================
# NOW START TEMPLATE PROCESSING
# ============================================================

import time

frames = []
upload_logs = []

# ============================================================
# Auto Process All Templates
# WITH TIMING
# ============================================================
import time
frames = []
upload_logs = []


for file in template_files:

    # ========================================================
    # Start Timer
    # ========================================================

    file_start_time = time.perf_counter()


    meta = parse_filename(
        file.name
    )


    # ========================================================
    # 1. TEMPLATE DETECTION
    # ========================================================

    detect_start_time = time.perf_counter()


    try:

        file.seek(0)

        detected_type_code = (
            detect_template_type(
                file
            )
        )


        detect_status = "SUCCESS"


    except Exception as detect_error:

        detected_type_code = "DETECTION_ERROR"

        detect_status = "FAILED"


    detect_seconds = (
        time.perf_counter()
        -
        detect_start_time
    )


    # ========================================================
    # If Detection Failed
    # ========================================================

    if (
        detected_type_code
        ==
        "DETECTION_ERROR"
    ):

        total_seconds = (
            time.perf_counter()
            -
            file_start_time
        )


        upload_logs.append({

            "檔案":
                file.name,

            "辨識格式":
                "DETECTION_ERROR",

            "Template Code":
                "DETECTION_ERROR",

            "狀態":
                "FAILED",

            "錯誤":
                str(
                    detect_error
                ),

            "原始筆數":
                0,

            "Detect 秒數":
                round(
                    detect_seconds,
                    3,
                ),

            "Transform 秒數":
                0,

            "總秒數":
                round(
                    total_seconds,
                    3,
                ),

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


        continue


    # ========================================================
    # 2. TRANSFORM
    # ========================================================

    transform_start_time = (
        time.perf_counter()
    )


    try:

        file.seek(0)


        (
            data,
            detected_template,
        ) = auto_transform_template(
            file,
            detected_type_code,
        )


        transform_seconds = (
            time.perf_counter()
            -
            transform_start_time
        )


        # ====================================================
        # Validate Result
        # ====================================================

        if data is None:

            raise ValueError(
                "Transformer 回傳 None。"
            )


        if not isinstance(
            data,
            pd.DataFrame,
        ):

            raise ValueError(
                "Transformer 沒有回傳 DataFrame。"
            )


        if data.empty:

            raise ValueError(
                "Transformer 回傳空 DataFrame。"
            )


        # ====================================================
        # Clean attrs
        # ====================================================

        data.attrs = {}


        # ====================================================
        # Save Successful Frame
        # ====================================================

        frames.append(
            data
        )


        # ====================================================
        # Total Time
        # ====================================================

        total_seconds = (
            time.perf_counter()
            -
            file_start_time
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

            # ================================================
            # TIMING
            # ================================================

            "Detect 秒數":
                round(
                    detect_seconds,
                    3,
                ),

            "Transform 秒數":
                round(
                    transform_seconds,
                    3,
                ),

            "總秒數":
                round(
                    total_seconds,
                    3,
                ),

            # ================================================
            # Filename Metadata
            # ================================================

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


    # ========================================================
    # Transform Failed
    # ========================================================

    except Exception as e:

        transform_seconds = (
            time.perf_counter()
            -
            transform_start_time
        )


        total_seconds = (
            time.perf_counter()
            -
            file_start_time
        )


        upload_logs.append({

            "檔案":
                file.name,

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

            # ================================================
            # TIMING
            # ================================================

            "Detect 秒數":
                round(
                    detect_seconds,
                    3,
                ),

            "Transform 秒數":
                round(
                    transform_seconds,
                    3,
                ),

            "總秒數":
                round(
                    total_seconds,
                    3,
                ),

            # ================================================
            # Filename Metadata
            # ================================================

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


# ============================================================
# Upload Log
# ============================================================

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
# STORE RECONCILIATION
# ============================================================

source_store_check = (
    detail
    .assign(
        QTY_CHECK=pd.to_numeric(
            detail["QTY"],
            errors="coerce",
        )
    )
    .groupby(
        [
            "YEAR",
            "MONTH",
            "FILE_OUTLET_NO",
        ],
        dropna=False,
        as_index=False,
    )
    .agg(
        SOURCE_ROWS=(
            "QTY_CHECK",
            "size",
        ),
        SOURCE_QTY=(
            "QTY_CHECK",
            "sum",
        ),
    )
)


source_store_check = (
    source_store_check
    .rename(
        columns={
            "YEAR": "年",
            "MONTH": "月",
            "FILE_OUTLET_NO": "店家編號",
        }
    )
)


output_store_check = (
    final_report
    .assign(
        QTY_CHECK=pd.to_numeric(
            final_report["銷量(瓶)"],
            errors="coerce",
        )
    )
    .groupby(
        [
            "年",
            "月",
            "店家編號",
        ],
        dropna=False,
        as_index=False,
    )
    .agg(
        店家名稱=(
            "店家名稱",
            "first",
        ),
        OUTPUT_ROWS=(
            "QTY_CHECK",
            "size",
        ),
        OUTPUT_QTY=(
            "QTY_CHECK",
            "sum",
        ),
    )
)


store_reconciliation = (
    source_store_check
    .merge(
        output_store_check,
        how="outer",
        on=[
            "年",
            "月",
            "店家編號",
        ],
    )
)


for column in [
    "SOURCE_ROWS",
    "OUTPUT_ROWS",
    "SOURCE_QTY",
    "OUTPUT_QTY",
]:

    store_reconciliation[column] = (
        pd.to_numeric(
            store_reconciliation[column],
            errors="coerce",
        )
        .fillna(0)
    )


store_reconciliation[
    "QTY_DIFF"
] = (
    store_reconciliation[
        "OUTPUT_QTY"
    ]
    -
    store_reconciliation[
        "SOURCE_QTY"
    ]
)


store_reconciliation[
    "MERGED_ROWS"
] = (
    store_reconciliation[
        "SOURCE_ROWS"
    ]
    -
    store_reconciliation[
        "OUTPUT_ROWS"
    ]
)


store_reconciliation[
    "STATUS"
] = (
    store_reconciliation[
        "QTY_DIFF"
    ]
    .apply(
        lambda x:
        "PASS"
        if abs(x) < 0.000001
        else "CHECK"
    )
)


# ============================================================
# TEMP DEBUG PREVIEW
# 先放這裡確認，之後再移到 Dashboard
# ============================================================

st.warning(
    "DEBUG：Store Reconciliation 已執行"
)

st.write(
    "store_reconciliation rows:",
    len(store_reconciliation)
)

st.dataframe(
    store_reconciliation,
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# Dashboard Metrics
# ============================================================

source_count = len(
    raw_all
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


# ============================================================
# Mapping Quality
# ============================================================

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


# ============================================================
# Responsible Sales / Customer Ownership
#
# 新邏輯：
# 不再檢查 FILE_SALES_ID 是否 MATCH
# 改成確認是否能透過 Customer Mapping
# 判定真正負責業務
# ============================================================
# ============================================================
# ============================================================
# Unable to Identify Responsible Sales
# FILE LEVEL
# ============================================================

unmapped_owner_files = (
    detail[
        detail[
            "CUSTOMER_OWNER_STATUS"
        ]
        .eq("UNMAPPED")
    ]
    [
        [
            "SOURCE_FILE",
            "FILE_OUTLET_NO",
            "FILE_RAWDATA_NAME",
        ]
    ]
    .drop_duplicates(
        subset=[
            "SOURCE_FILE"
        ]
    )
    .reset_index(
        drop=True
    )
)


customer_owner_unmapped = len(
    unmapped_owner_files
)

# ============================================================
# Audit Metrics
#
# Contract mismatch 只做 Audit
# 不影響 Submission matching
# ============================================================

contract_mismatch = int(
    detail[
        "CONTRACT_CHECK"
    ]
    .eq("MISMATCH")
    .sum()
)


# ============================================================
# File Processing Status
# ============================================================

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
        "🐄 Source Count (來源資料)",
        f"{source_count:,}",
    )


with d2:

    st.metric(
        "🧀 Output Count (輸出資料)",
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
        "✔️ Successful Files (成功檔案)",
        f"{successful_files:,}",
    )


with d6:

    st.metric(
        "❌ Failed Files (失敗檔案)",
        f"{failed_files:,}",
    )

# ============================================================
# ============================================================
# STORE-LEVEL RECONCILIATION
# ============================================================

store_pass_count = int(
    store_reconciliation[
        "STATUS"
    ]
    .eq("PASS")
    .sum()
)


store_check_count = int(
    store_reconciliation[
        "STATUS"
    ]
    .eq("CHECK")
    .sum()
)


total_store_count = len(
    store_reconciliation
)


# ========================================================
# KPI
# ========================================================

r1, r2, r3 = st.columns(3)


with r1:

    st.metric(
        "🏪 Total Store Count (核對店家)",
        f"{total_store_count:,}",
    )


with r2:

    st.metric(
        "✅ PASS",
        f"{store_pass_count:,}",
    )


with r3:

    st.metric(
        "⚠️ CHECK",
        f"{store_check_count:,}",
    )


# ========================================================
# Result
# ========================================================

if store_check_count == 0:

    st.success(
        "✅ 所有店家銷量核對完成："
        "來源資料與最終輸出總銷量一致。"
        "Store check count have been validated."
    )


else:

    st.error(
        f"⚠️ 有 {store_check_count:,} 間店家 "
        "來源與輸出銷量不一致。"
    )


    reconciliation_issues = (
        store_reconciliation[
            store_reconciliation[
                "STATUS"
            ]
            .eq("CHECK")
        ]
        .copy()
    )


    st.dataframe(
        reconciliation_issues,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# ============================================================
# Additional Quality Metrics
# ============================================================

q1, q2, q3 = st.columns(3)

with q1:
    st.metric(
        "Customer + SKU 同時 Unmapped",
        f"{both_unmapped:,}",
    )

with q2:
    st.metric(
        "無法判定負責業務（檔案）",
        f"{customer_owner_unmapped:,}",
    )

with q3:
    st.metric(
        "合約編號不一致（Audit）",
        f"{contract_mismatch:,}",
    )


# ============================================================
# Unmapped Responsible Sales - File Detail
# ============================================================

if customer_owner_unmapped > 0:

    with st.expander(
        "查看無法判定負責業務的檔案",
        expanded=False,
    ):

        st.dataframe(
            unmapped_owner_files,
            width="stretch",
            hide_index=True,
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
#
# IMPORTANT:
# 業務歸屬直接使用 Customer Mapping 裡的：
# Sales ID / Sales
#
# Submission matching 不再使用：
# - File Sales ID
# - Contract JDE
#
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
#
# 沒有 Customer ID / Outlet No 的資料
# 無法作為應繳店家
# ============================================================

expected_submission = (
    expected_submission[
        expected_submission[
            "Outlet No"
        ]
        .map(
            norm_id_key
        )
        .ne("")
    ]
    .copy()
)


# ============================================================
# 8. EXPECTED KEYS
# ============================================================

# ------------------------------------------------------------
# Reporting Period
# ------------------------------------------------------------

expected_submission[
    "PERIOD_KEY"
] = selected_period


# ------------------------------------------------------------
# Customer / Outlet ID
#
# 這是新的主要 tracking key
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# Responsible Sales
#
# 只供 Dashboard / 業務繳交進度使用
# 不參與 EXPECTED_KEY matching
# ------------------------------------------------------------

expected_submission[
    "RESPONSIBLE_SALES_ID"
] = (
    expected_submission[
        "Sales ID"
    ]
    .map(
        norm_id_key
    )
)


expected_submission[
    "RESPONSIBLE_SALES"
] = (
    expected_submission[
        "Sales"
    ]
    .fillna("")
    .map(
        norm_text
    )
)


# ============================================================
# 9. EXPECTED KEY
#
# NEW RULE:
#
# YYYYMM | Customer ID
#
# Example:
# 202609|32014065
#
# 不再使用：
# Sales ID
# Contract JDE
# ============================================================

expected_submission[
    "EXPECTED_KEY"
] = (
    expected_submission[
        "PERIOD_KEY"
    ]
    + "|"
    + expected_submission[
        "OUTLET_TRACKING_KEY"
    ]
)


# ============================================================
# 10. EXPECTED DUPLICATE REMOVAL
#
# 同月份 + 同 Customer ID
# 只算一個應繳店家
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


# ============================================================
# Export Data
# ============================================================

st.divider()


# ============================================================
# STEP 5 - EXPORT
# ============================================================

section_header(
    "5",
    "Export",
    "Download final report, audit records and unmapped records",
)


# ============================================================
# 1. EXPORT PERIOD
#
# 用 selected_period 當檔名
# 如果沒有 selected_period，就用 ALL
# ============================================================

export_period = (
    str(selected_period)
    if (
        "selected_period" in locals()
        and selected_period
    )
    else "ALL"
)


# ============================================================
# 2. SAFE COLUMN HELPER
#
# 避免某些 Template 沒有特定 Audit column
# 造成 KeyError
# ============================================================

def existing_columns(
    dataframe,
    columns,
):

    return [
        column
        for column in columns
        if column in dataframe.columns
    ]


# ============================================================
# 3. DETAIL-LEVEL AUDIT COLUMNS
#
# IMPORTANT:
#
# detail = 原始 row-level
# final_report = aggregation 後資料
#
# 所以兩者不再直接 concat
# ============================================================

audit_export_columns = [

    # --------------------------------------------------------
    # Source
    # --------------------------------------------------------
    "SOURCE_FILE",
    "SOURCE_SHEET",
    "SOURCE_ROW",
    "TEMPLATE_TYPE",

    # --------------------------------------------------------
    # Filename Metadata
    # --------------------------------------------------------
    "FILE_PERIOD",
    "FILE_SALES_ID",
    "FILE_SALES_ID_AUDIT",
    "FILE_CONTRACT_JDE",
    "FILE_OUTLET_NO",
    "FILE_RAWDATA_NAME",

    # --------------------------------------------------------
    # Customer Ownership
    # --------------------------------------------------------
    "SUBMISSION_CUSTOMER_ID",
    "RESPONSIBLE_SALES_ID",
    "RESPONSIBLE_SALES",
    "RESPONSIBLE_SALES_STATUS",
    "CUSTOMER_OWNER_STATUS",

    # --------------------------------------------------------
    # Customer Mapping Result
    # --------------------------------------------------------
    "Sales ID",
    "Sales",
    "Contract JDE",
    "Contract NAME",
    "Outlet No",
    "Outlet NAME",
    "Rawdata Name",

    # --------------------------------------------------------
    # Customer Mapping Status
    # --------------------------------------------------------
    "CUSTOMER_MAPPING_STATUS",
    "OUTLET_CHECK",
    "CONTRACT_CHECK",
    "RAWDATA_NAME_CHECK",

    # --------------------------------------------------------
    # Raw Product
    # --------------------------------------------------------
    "RAW_PRODUCT_CODE",
    "RAW_SKU",
    "RAW_BRAND",
    "RAW_SIZE",
    "RAW_CATEGORY",
    "QTY",

    # --------------------------------------------------------
    # SKU Mapping
    # --------------------------------------------------------
    "SKU_MAPPING_STATUS",
    "Mapping Name",
    "RSP",
    "Price Band",
    "Manufacture",
    "Brand",
    "SKU Name",
    "SIZE",
    "CATEGORY",

    # --------------------------------------------------------
    # Price Audit
    # --------------------------------------------------------
    "RAW_UNIT_PRICE",
    "RAW_TOTAL_PRICE",
    "UNIT_PRICE",
    "TOTAL_PRICE",

    # --------------------------------------------------------
    # Detection Audit
    # --------------------------------------------------------
    "SOURCE_QTY_COLUMN",
    "SOURCE_PRICE_COLUMN",
]


audit_export_columns = existing_columns(
    detail,
    audit_export_columns,
)


# ============================================================
# 4. COMPLETE AUDIT DETAIL
#
# Row-level
# 不跟 final_report 橫向拼接
# ============================================================

audit_export = (
    detail[
        audit_export_columns
    ]
    .copy()
    .reset_index(
        drop=True
    )
)


# ============================================================
# 5. CUSTOMER UNMAPPED
# ============================================================

customer_unmapped_mask = (
    detail[
        "CUSTOMER_MAPPING_STATUS"
    ]
    .eq("UNMAPPED")
)


customer_unmapped_export = (
    detail.loc[
        customer_unmapped_mask,
        audit_export_columns,
    ]
    .copy()
    .reset_index(
        drop=True
    )
)


# ============================================================
# 6. SKU UNMAPPED
# ============================================================

sku_unmapped_mask = (
    detail[
        "SKU_MAPPING_STATUS"
    ]
    .eq("UNMAPPED")
)


sku_unmapped_export = (
    detail.loc[
        sku_unmapped_mask,
        audit_export_columns,
    ]
    .copy()
    .reset_index(
        drop=True
    )
)


# ============================================================
# 7. ALL UNMAPPED
#
# Customer OR SKU Unmapped
# ============================================================

unmapped_mask = (
    customer_unmapped_mask
    |
    sku_unmapped_mask
)


unmapped = (
    detail.loc[
        unmapped_mask,
        audit_export_columns,
    ]
    .copy()
    .reset_index(
        drop=True
    )
)


# ============================================================
# 8. CUSTOMER + SKU BOTH UNMAPPED
# ============================================================

both_unmapped_mask = (
    detail[
        "CUSTOMER_MAPPING_STATUS"
    ]
    .eq("UNMAPPED")
    &
    detail[
        "SKU_MAPPING_STATUS"
    ]
    .eq("UNMAPPED")
)


both_unmapped_export = (
    detail.loc[
        both_unmapped_mask,
        audit_export_columns,
    ]
    .copy()
    .reset_index(
        drop=True
    )
)


# ============================================================
# 9. RESPONSIBLE SALES / OWNERSHIP AUDIT
#
# FILE LEVEL
#
# 一個 Source File 只保留一筆
# ============================================================

ownership_columns = existing_columns(
    detail,
    [
        "SOURCE_FILE",

        # Filename
        "FILE_PERIOD",
        "FILE_SALES_ID_AUDIT",
        "FILE_OUTLET_NO",
        "FILE_CONTRACT_JDE",
        "FILE_RAWDATA_NAME",

        # Mapping
        "Outlet No",
        "Outlet NAME",

        # True Responsible Sales
        "RESPONSIBLE_SALES_ID",
        "RESPONSIBLE_SALES",

        # Status
        "CUSTOMER_MAPPING_STATUS",
        "RESPONSIBLE_SALES_STATUS",
        "CUSTOMER_OWNER_STATUS",

        # Audit
        "OUTLET_CHECK",
        "CONTRACT_CHECK",
        "RAWDATA_NAME_CHECK",
    ],
)


ownership_audit = (
    detail[
        ownership_columns
    ]
    .copy()
)


if "SOURCE_FILE" in ownership_audit.columns:

    ownership_audit = (
        ownership_audit
        .drop_duplicates(
            subset=[
                "SOURCE_FILE"
            ],
            keep="first",
        )
        .reset_index(
            drop=True
        )
    )

else:

    ownership_audit = (
        ownership_audit
        .drop_duplicates()
        .reset_index(
            drop=True
        )
    )


# ============================================================
# 10. RENAME OWNERSHIP AUDIT
#
# 下載 Excel 比較容易看懂
# ============================================================

ownership_audit = (
    ownership_audit
    .rename(
        columns={

            "SOURCE_FILE":
                "檔案",

            "FILE_PERIOD":
                "年月",

            "FILE_SALES_ID_AUDIT":
                "檔名業務代號",

            "FILE_OUTLET_NO":
                "檔名Customer ID",

            "FILE_CONTRACT_JDE":
                "檔名合約編號",

            "FILE_RAWDATA_NAME":
                "檔名店家名稱",

            "Outlet No":
                "Mapping Customer ID",

            "Outlet NAME":
                "Mapping 店家名稱",

            "RESPONSIBLE_SALES_ID":
                "負責業務代號",

            "RESPONSIBLE_SALES":
                "負責業務",

            "CUSTOMER_MAPPING_STATUS":
                "Customer Mapping",

            "RESPONSIBLE_SALES_STATUS":
                "負責業務 Mapping",

            "CUSTOMER_OWNER_STATUS":
                "業務歸屬狀態",

            "OUTLET_CHECK":
                "Customer ID Check",

            "CONTRACT_CHECK":
                "Contract Check",

            "RAWDATA_NAME_CHECK":
                "店名 Check",
        }
    )
)


# ============================================================
# 11. UNMAPPED RESPONSIBLE SALES
#
# FILE LEVEL
# ============================================================

if "業務歸屬狀態" in ownership_audit.columns:

    unmapped_owner_export = (
        ownership_audit[
            ownership_audit[
                "業務歸屬狀態"
            ]
            .eq("UNMAPPED")
        ]
        .copy()
        .reset_index(
            drop=True
        )
    )

else:

    unmapped_owner_export = pd.DataFrame()


# ============================================================
# 12. SUBMISSION TRACKER
#
# 現在真正的 submission tracking dataframe
# 是 submission_detail
#
# 不再使用舊的 submission 變數
# ============================================================

if (
    "submission_detail"
    in locals()
    and
    isinstance(
        submission_detail,
        pd.DataFrame,
    )
):

    submission_export = (
        submission_detail
        .copy()
        .reset_index(
            drop=True
        )
    )

else:

    submission_export = pd.DataFrame()


# ============================================================
# 13. SALES SUBMISSION SUMMARY
# ============================================================

if (
    "sales_submission"
    in locals()
    and
    isinstance(
        sales_submission,
        pd.DataFrame,
    )
):

    sales_submission_export = (
        sales_submission
        .copy()
        .reset_index(
            drop=True
        )
    )

else:

    sales_submission_export = (
        pd.DataFrame()
    )


# ============================================================
# 14. UNMATCHED UPLOAD
#
# 有上傳，但 Customer ID 無法對 Customer Mapping
# ============================================================

if (
    "unmatched_uploads"
    in locals()
    and
    isinstance(
        unmatched_uploads,
        pd.DataFrame,
    )
):

    unmatched_upload_export = (
        unmatched_uploads
        .copy()
        .reset_index(
            drop=True
        )
    )

else:

    unmatched_upload_export = (
        pd.DataFrame()
    )


# ============================================================
# 15. STORE RECONCILIATION
# ============================================================

if (
    "store_reconciliation"
    in locals()
    and
    isinstance(
        store_reconciliation,
        pd.DataFrame,
    )
):

    store_reconciliation_export = (
        store_reconciliation
        .copy()
        .reset_index(
            drop=True
        )
    )

else:

    store_reconciliation_export = (
        pd.DataFrame()
    )


# ============================================================
# 16. CLEAN DATAFRAME ATTRS
#
# 避免 Excel / Arrow serialization 問題
# ============================================================

export_dataframes = [
    final_report,
    upload_log,
    submission_export,
    sales_submission_export,
    unmapped,
    customer_unmapped_export,
    sku_unmapped_export,
    both_unmapped_export,
    audit_export,
    ownership_audit,
    unmapped_owner_export,
    unmatched_upload_export,
    customer_duplicates,
    sku_duplicates,
    store_reconciliation_export,
]


for dataframe in export_dataframes:

    if isinstance(
        dataframe,
        pd.DataFrame,
    ):

        dataframe.attrs = {}


# ============================================================
# 17. FULL WORKBOOK
# ============================================================

workbook_sheets = {

    # --------------------------------------------------------
    # Main Report
    # --------------------------------------------------------
    "Final Report":
        final_report,

    # --------------------------------------------------------
    # File Processing
    # --------------------------------------------------------
    "Upload Log":
        upload_log,

    # --------------------------------------------------------
    # Submission Tracking
    # --------------------------------------------------------
    "Submission Tracker":
        submission_export,

    "Sales Submission":
        sales_submission_export,

    # --------------------------------------------------------
    # Mapping Problems
    # --------------------------------------------------------
    "All Unmapped":
        unmapped,

    "Customer Unmapped":
        customer_unmapped_export,

    "SKU Unmapped":
        sku_unmapped_export,

    "Both Unmapped":
        both_unmapped_export,

    # --------------------------------------------------------
    # Audit
    # --------------------------------------------------------
    "Audit Detail":
        audit_export,

    "Sales Ownership Audit":
        ownership_audit,

    "Owner Unmapped":
        unmapped_owner_export,

    "Unmatched Upload":
        unmatched_upload_export,

    # --------------------------------------------------------
    # Duplicate Mapping
    # --------------------------------------------------------
    "Customer Duplicates":
        customer_duplicates,

    "SKU Duplicates":
        sku_duplicates,

    # --------------------------------------------------------
    # Reconciliation
    # --------------------------------------------------------
    "Store Reconciliation":
        store_reconciliation_export,
}


# ============================================================
# 18. REMOVE EMPTY OPTIONAL SHEETS
#
# Final Report / Upload Log 保留
# 其他完全空白的 sheet 可以拿掉
# ============================================================

required_sheet_names = {
    "Final Report",
    "Upload Log",
}


workbook_sheets = {

    sheet_name:
        dataframe

    for (
        sheet_name,
        dataframe
    )
    in workbook_sheets.items()

    if (
        sheet_name
        in required_sheet_names
        or
        (
            isinstance(
                dataframe,
                pd.DataFrame,
            )
            and
            not dataframe.empty
        )
    )
}


# ============================================================
# 19. GENERATE FULL EXCEL
# ============================================================

workbook = excel_bytes(
    workbook_sheets
)


# ============================================================
# 20. FINAL CSV
# ============================================================

final_csv = (
    final_report
    .to_csv(
        index=False
    )
    .encode(
        "utf-8-sig"
    )
)


# ============================================================
# 21. UNMAPPED WORKBOOK
#
# 獨立 Unmapped Excel
# ============================================================

unmapped_workbook_sheets = {

    "All Unmapped":
        unmapped,

    "Customer Unmapped":
        customer_unmapped_export,

    "SKU Unmapped":
        sku_unmapped_export,

    "Both Unmapped":
        both_unmapped_export,

    "Owner Unmapped":
        unmapped_owner_export,
}


unmapped_workbook_sheets = {

    sheet_name:
        dataframe

    for (
        sheet_name,
        dataframe
    )
    in unmapped_workbook_sheets.items()

    if (
        isinstance(
            dataframe,
            pd.DataFrame,
        )
        and
        not dataframe.empty
    )
}


# 如果完全沒有 Unmapped
if not unmapped_workbook_sheets:

    unmapped_workbook_sheets = {

        "Unmapped":
            pd.DataFrame({
                "Status": [
                    "No unmapped records"
                ]
            })
    }


unmapped_workbook = excel_bytes(
    unmapped_workbook_sheets
)


# ============================================================
# 22. EXPORT SUMMARY
# ============================================================

export_summary_col1, export_summary_col2, export_summary_col3 = (
    st.columns(3)
)


with export_summary_col1:

    st.metric(
        "Final Report",
        f"{len(final_report):,}",
    )


with export_summary_col2:

    st.metric(
        "Unmapped Rows",
        f"{len(unmapped):,}",
    )


with export_summary_col3:

    st.metric(
        "Owner Unmapped Files",
        f"{len(unmapped_owner_export):,}",
    )


# ============================================================
# 23. DOWNLOAD BUTTONS
# ============================================================

export_col1, export_col2, export_col3 = (
    st.columns(3)
)


# ------------------------------------------------------------
# Full Excel
# ------------------------------------------------------------

with export_col1:

    st.download_button(
        "📥 下載完整 Excel",
        data=workbook,
        file_name=(
            f"POS_Monthly_Report_"
            f"{export_period}.xlsx"
        ),
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        width="stretch",

        # Streamlit 1.66+
        # Download 時不要重新跑整支程式
        on_click="ignore",
    )


# ------------------------------------------------------------
# Final CSV
# ------------------------------------------------------------

with export_col2:

    st.download_button(
        "📄 下載 Final CSV",
        data=final_csv,
        file_name=(
            f"POS_Monthly_Final_"
            f"{export_period}.csv"
        ),
        mime="text/csv",
        width="stretch",
        on_click="ignore",
    )


# ------------------------------------------------------------
# Unmapped Excel
# ------------------------------------------------------------

with export_col3:

    st.download_button(
        "⚠️ 下載 Unmapped Excel",
        data=unmapped_workbook,
        file_name=(
            f"POS_Unmapped_"
            f"{export_period}.xlsx"
        ),
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        width="stretch",
        on_click="ignore",
    )

# ============================================================
# Protection Notes
# ============================================================
st.subheader(
    "🏪 All Store Reconciliation"
)

st.dataframe(
    store_reconciliation,
    use_container_width=True,
    hide_index=True,
)


reconciliation_issues = (
    store_reconciliation[
        store_reconciliation[
            "STATUS"
        ]
        .eq("CHECK")
    ]
    .copy()
)


st.subheader(
    "⚠️ CHECK Only"
)

st.dataframe(
    reconciliation_issues,
    use_container_width=True,
    hide_index=True,
)

