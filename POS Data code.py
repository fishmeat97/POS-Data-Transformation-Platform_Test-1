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


def numeric(series):
    if series is None:
        return pd.Series(dtype="float64")

    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("NT$", "", regex=False)
        .str.replace("NTD", "", regex=False)
        .str.strip(),
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
        ["貨品名稱"],
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
# Auto Detect Template Type
# ============================================================
def detect_template_type(uploaded_file):
    suffix = Path(
        uploaded_file.name
    ).suffix.lower()

    # --------------------------------------------------------
    # CSV -> Template 1
    # --------------------------------------------------------
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
                ],
            )

            qty_col = first_matching_column(
                df,
                [
                    "銷量 Quantity",
                    "Quantity",
                    "銷量",
                    "數量",
                ],
            )

            if (
                sku_col is not None
                and qty_col is not None
            ):
                return "TEMPLATE_1"

        except Exception:
            pass

    # --------------------------------------------------------
    # Excel
    # --------------------------------------------------------
    if suffix in [
        ".xlsx",
        ".xls",
    ]:
        uploaded_file.seek(0)

        excel_file = pd.ExcelFile(
            uploaded_file
        )

        # ----------------------------------------------------
        # Multi-sheet 國泰 style
        # ----------------------------------------------------
        if len(excel_file.sheet_names) > 1:
            try:
                uploaded_file.seek(0)

                first_sheet = pd.read_excel(
                    uploaded_file,
                    sheet_name=excel_file.sheet_names[0],
                    header=0,
                )

                columns = {
                    norm_col(c)
                    for c in first_sheet.columns
                }

                if any(
                    "品名規格" in c
                    for c in columns
                ):
                    return "TEMPLATE_3"

            except Exception:
                pass

        # ----------------------------------------------------
        # Read first sheet as raw
        # ----------------------------------------------------
        uploaded_file.seek(0)

        raw = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            header=None,
        )

        # ----------------------------------------------------
        # Template 2
        # ----------------------------------------------------
        for i in range(
            min(
                20,
                len(raw),
            )
        ):
            values = [
                norm_text(x)
                for x in raw.iloc[i].tolist()
            ]

            if (
                "商品代號" in values
                and "商品名稱" in values
                and any(
                    "實銷" in v
                    for v in values
                )
            ):
                return "TEMPLATE_2"

        # ----------------------------------------------------
        # Template 3 single-sheet
        # ----------------------------------------------------
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
                "數" in v
                and "量" in v
                for v in values
            )

            if (
                has_product
                and has_qty
            ):
                return "TEMPLATE_3"

        # ----------------------------------------------------
        # Template 1 Excel version
        # ----------------------------------------------------
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
                ],
            )

            qty_col = first_matching_column(
                df,
                [
                    "銷量 Quantity",
                    "Quantity",
                    "銷量",
                    "數量",
                ],
            )

            if (
                sku_col is not None
                and qty_col is not None
            ):
                return "TEMPLATE_1"

        except Exception:
            pass

    return "UNKNOWN"


# ============================================================
# Unified Auto Transformer
# ============================================================
def auto_transform_template(uploaded_file):
    uploaded_file.seek(0)

    template_type = detect_template_type(
        uploaded_file
    )

    uploaded_file.seek(0)

    if template_type == "TEMPLATE_1":
        result = transform_template_1(
            uploaded_file
        )

        return (
            result,
            "Template 1｜標準 POS SKU 表",
        )

    elif template_type == "TEMPLATE_2":
        result = transform_template_2(
            uploaded_file
        )

        return (
            result,
            "Template 2｜月份實銷表",
        )

    elif template_type == "TEMPLATE_3":
        uploaded_file.seek(0)

        excel_file = pd.ExcelFile(
            uploaded_file
        )

        # Multi-sheet first
        if len(excel_file.sheet_names) > 1:
            try:
                uploaded_file.seek(0)

                result = transform_template_3_multisheet(
                    uploaded_file
                )

                return (
                    result,
                    "Template 3A｜多工作表銷售明細",
                )

            except Exception:
                pass

        uploaded_file.seek(0)

        result = transform_template_3_salesdetail(
            uploaded_file
        )

        return (
            result,
            "Template 3B｜單工作表銷貨明細",
        )

    raise ValueError(
        "無法辨識此檔案的 Template 格式"
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

    # ========================================================
    # Customer Mapping Status
    # ========================================================

    merged[
        "CUSTOMER_MAPPING_STATUS"
    ] = merged[
        "Outlet No"
    ].apply(
        lambda x:
        "MAPPED"
        if norm_text(x)
        else "UNMAPPED"
    )


    # ========================================================
    # Sales ID Check
    # ========================================================

    merged[
        "SALES_ID_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if (
                row["CUSTOMER_MAPPING_STATUS"]
                == "UNMAPPED"
            )
            else (
                "MATCH"
                if (
                    norm_key(
                        row["FILE_SALES_ID"]
                    )
                    ==
                    norm_key(
                        row["Sales ID"]
                    )
                )
                else "MISMATCH"
            )
        ),
        axis=1,
    )


    # ========================================================
    # Contract Check
    # ========================================================

    merged[
        "CONTRACT_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if (
                row["CUSTOMER_MAPPING_STATUS"]
                == "UNMAPPED"
            )
            else (
                "MATCH"
                if (
                    norm_key(
                        row["FILE_CONTRACT_JDE"]
                    )
                    ==
                    norm_key(
                        row["Contract JDE"]
                    )
                )
                else "MISMATCH"
            )
        ),
        axis=1,
    )


    # ========================================================
    # Outlet No Check
    # ========================================================

    merged[
        "OUTLET_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if (
                row["CUSTOMER_MAPPING_STATUS"]
                == "UNMAPPED"
            )
            else (
                "MATCH"
                if (
                    norm_key(
                        row["FILE_OUTLET_NO"]
                    )
                    ==
                    norm_key(
                        row["Outlet No"]
                    )
                )
                else "MISMATCH"
            )
        ),
        axis=1,
    )


    # ========================================================
    # Rawdata Name Check
    # ========================================================

    merged[
        "RAWDATA_NAME_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if (
                row["CUSTOMER_MAPPING_STATUS"]
                == "UNMAPPED"
            )
            else (
                "MATCH"
                if (
                    norm_key(
                        row["FILE_RAWDATA_NAME"]
                    )
                    ==
                    norm_key(
                        row["Rawdata Name"]
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
        "Outlet No",
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

    sku_file = st.file_uploader(
        "SKU Mapping",
        type=[
            "xlsx",
            "xls",
            "csv",
        ],
        key="sku",
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
    or sku_file is None
):
    st.info(
        "請上傳 Customer Mapping 與 SKU Mapping。"
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
# Load SKU Mapping
# ============================================================
try:
    sku_raw = read_mapping_file(
        sku_file,
        preferred_sheet="SKU data",
    )

    (
        sku_map,
        sku_duplicates,
    ) = prepare_sku_mapping(
        sku_raw
    )

except Exception as e:
    st.error(
        f"SKU Mapping 讀取失敗：{e}"
    )
    st.stop()


# ============================================================
# Auto Process All Templates
# ============================================================
frames = []
upload_logs = []

for file in template_files:
    meta = parse_filename(
        file.name
    )

    try:
        file.seek(0)

        (
            data,
            detected_template,
        ) = auto_transform_template(
            file
        )

        frames.append(
            data
        )

        upload_logs.append({
            "檔案":
                file.name,

            "辨識格式":
                detected_template,

            "狀態":
                "SUCCESS",

            "錯誤":
                "",

            "原始筆數":
                len(data),

            "業務代號":
                meta["FILE_SALES_ID"],

            "合約編號":
                meta["FILE_CONTRACT_JDE"],

            "年月":
                meta["FILE_PERIOD"],

            "店家編號":
                meta["FILE_OUTLET_NO"],

            "分店/Rawdata Name":
                meta["FILE_RAWDATA_NAME"],

            "檔名格式":
                meta["FILENAME_STATUS"],
        })

    except Exception as e:
        upload_logs.append({
            "檔案":
                file.name,

            "辨識格式":
                "UNKNOWN",

            "狀態":
                "FAILED",

            "錯誤":
                str(e),

            "原始筆數":
                0,

            "業務代號":
                meta["FILE_SALES_ID"],

            "合約編號":
                meta["FILE_CONTRACT_JDE"],

            "年月":
                meta["FILE_PERIOD"],

            "店家編號":
                meta["FILE_OUTLET_NO"],

            "分店/Rawdata Name":
                meta["FILE_RAWDATA_NAME"],

            "檔名格式":
                meta["FILENAME_STATUS"],
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
# ============================================================
# STEP 3 - DASHBOARD
# ============================================================

section_header(
    "3",
    "Dashboard",
    "Data volume, mapping quality and file processing status",
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
        "📋 本月應繳",
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
    "Customer Mapping 重複",
    "SKU Mapping 重複",
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
    "Export",
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
