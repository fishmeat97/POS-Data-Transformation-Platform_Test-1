import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# Page
# ============================================================
st.set_page_config(
    page_title="POS Monthly Data Mapping",
    page_icon="📊",
    layout="wide",
)

st.title("📊 POS Monthly Data Mapping")
st.caption(
    "三種 Template 分開上傳，自動轉換成同一個月報格式，"
    "再執行 Customer / SKU Mapping 與資料勾稽。"
)


# ============================================================
# Final output columns
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
# Customer Mapping Required Fields
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


# ============================================================
# SKU Mapping Required Fields
# ============================================================
SKU_REQUIRED = [
    "Mapping Name",
    "Manufacture",
    "Band",
    "RSP",
    "Price Band",
    "SIZE",
    "CATEGORY",
]


# ============================================================
# Helper Functions
# ============================================================
def norm_text(value):

    if pd.isna(value):
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(value)
    )

    text = text.strip()

    text = re.sub(
        r"\s+",
        " ",
        text
    )

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

    text = re.sub(
        r"[_\-]+",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text


# ============================================================
# Remove duplicated "(1)" from filename
# ============================================================
def clean_file_label(filename):

    stem = Path(filename).stem

    stem = re.sub(
        r"\(\d+\)$",
        "",
        stem
    ).strip()

    return stem


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

    sales_id = (
        parts[0]
        if len(parts) > 0
        else ""
    )

    contract_jde = (
        parts[1]
        if len(parts) > 1
        else ""
    )

    ym = (
        parts[2]
        if len(parts) > 2
        else ""
    )

    outlet_no = (
        parts[3]
        if len(parts) > 3
        else ""
    )

    rawdata_name = (
        "_".join(parts[4:])
        if len(parts) > 4
        else ""
    )


    # --------------------------------------------------------
    # Parse Year Month
    # --------------------------------------------------------
    year = None
    month = None

    digits = re.sub(
        r"\D",
        "",
        ym
    )


    # YYYYMM
    if len(digits) == 6:

        y = int(digits[:4])
        m = int(digits[4:])

        if (
            1900 <= y <= 2200
            and 1 <= m <= 12
        ):

            year = y
            month = m


    # ROC Year: 11509
    elif len(digits) == 5:

        y = (
            int(digits[:3])
            + 1911
        )

        m = int(
            digits[3:]
        )

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

        "FILE_SALES_ID":
            sales_id,

        "FILE_CONTRACT_JDE":
            contract_jde,

        "FILE_PERIOD":
            ym,

        "FILE_OUTLET_NO":
            outlet_no,

        "FILE_RAWDATA_NAME":
            rawdata_name,

        "YEAR":
            year,

        "MONTH":
            month,

        "FILENAME_STATUS":
            "OK"
            if ok
            else "CHECK",
    }


# ============================================================
# Flexible CSV Reader
# ============================================================
def read_csv_flexible(
    uploaded_file,
    header=0
):

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


# ============================================================
# Excel Reader
# ============================================================
def read_excel_bytes(
    uploaded_file,
    **kwargs
):

    uploaded_file.seek(0)

    return pd.read_excel(
        uploaded_file,
        **kwargs
    )


# ============================================================
# Find Matching Column
# ============================================================
def first_matching_column(
    df,
    keywords
):

    columns = list(
        df.columns
    )


    # Exact match
    for column in columns:

        current = norm_col(
            column
        )

        for keyword in keywords:

            if norm_col(
                keyword
            ) == current:

                return column


    # Contains match
    for column in columns:

        current = norm_col(
            column
        )

        for keyword in keywords:

            keyword_normalized = (
                norm_col(
                    keyword
                )
            )

            if (
                keyword_normalized
                and keyword_normalized
                in current
            ):

                return column


    return None


# ============================================================
# Convert to numeric
# ============================================================
def numeric(series):

    return pd.to_numeric(

        series.astype(str)

        .str.replace(
            ",",
            "",
            regex=False
        )

        .str.replace(
            "$",
            "",
            regex=False
        )

        .str.replace(
            "NT$",
            "",
            regex=False
        )

        .str.strip(),

        errors="coerce",
    )


# ============================================================
# Build basic records
# ============================================================
def base_records(
    df,
    meta,
    template_type,
    source_file
):

    output = pd.DataFrame(
        index=df.index
    )

    output[
        "SOURCE_FILE"
    ] = source_file

    output[
        "SOURCE_ROW"
    ] = range(
        1,
        len(df) + 1
    )

    output[
        "TEMPLATE_TYPE"
    ] = template_type


    for key, value in (
        meta.items()
    ):

        output[
            key
        ] = value


    return output


# ============================================================
# TEMPLATE 1
#
# Standard POS SKU Table
# Example: 冠德 CSV
# ============================================================
def transform_template_1(
    uploaded_file
):

    meta = parse_filename(
        uploaded_file.name
    )


    suffix = Path(
        uploaded_file.name
    ).suffix.lower()


    if suffix == ".csv":

        df = read_csv_flexible(
            uploaded_file
        )

    else:

        df = read_excel_bytes(
            uploaded_file,
            header=0
        )


    df = (
        df
        .dropna(
            how="all"
        )
        .reset_index(
            drop=True
        )
    )


    # --------------------------------------------------------
    # Detect SKU
    # --------------------------------------------------------
    sku_col = (
        first_matching_column(
            df,
            [
                "品項名稱 SKU Name",
                "SKU Name",
                "品項名稱",
                "品名",
                "商品名稱",
            ]
        )
    )


    # --------------------------------------------------------
    # Detect Quantity
    # --------------------------------------------------------
    qty_col = (
        first_matching_column(
            df,
            [
                "銷量 Quantity",
                "Quantity",
                "銷量",
                "數量",
                "瓶數",
            ]
        )
    )


    if (
        sku_col is None
        or qty_col is None
    ):

        raise ValueError(

            "Template 1 找不到必要欄位。"
            f" SKU={sku_col},"
            f" Quantity={qty_col}"
        )


    output = base_records(

        df,
        meta,
        "T1_STANDARD_POS",
        uploaded_file.name,
    )


    # --------------------------------------------------------
    # Customer
    # --------------------------------------------------------
    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    # --------------------------------------------------------
    # SKU
    # --------------------------------------------------------
    output[
        "RAW_SKU"
    ] = df[
        sku_col
    ].map(
        norm_text
    )


    # --------------------------------------------------------
    # Quantity
    # --------------------------------------------------------
    output[
        "QTY"
    ] = numeric(
        df[
            qty_col
        ]
    )


    # --------------------------------------------------------
    # Price
    # --------------------------------------------------------
    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    # --------------------------------------------------------
    # Optional Fields
    # --------------------------------------------------------
    brand_col = (
        first_matching_column(
            df,
            [
                "品牌名稱 Brand Name",
                "品牌名稱",
                "Brand Name",
            ]
        )
    )


    size_col = (
        first_matching_column(
            df,
            [
                "規格 Size (ml)",
                "Size (ml)",
                "規格",
                "容量",
            ]
        )
    )


    category_col = (
        first_matching_column(
            df,
            [
                "品類 Category",
                "Category",
                "品類",
            ]
        )
    )


    output[
        "RAW_BRAND"
    ] = (
        df[
            brand_col
        ]
        if brand_col
        else ""
    )


    output[
        "RAW_SIZE"
    ] = (
        df[
            size_col
        ]
        if size_col
        else ""
    )


    output[
        "RAW_CATEGORY"
    ] = (
        df[
            category_col
        ]
        if category_col
        else ""
    )


    return output


# ============================================================
# TEMPLATE 2
#
# Horizontal Monthly Actual Sales
# Example: 洋酒城
# ============================================================
def transform_template_2(
    uploaded_file
):

    meta = parse_filename(
        uploaded_file.name
    )


    raw = read_excel_bytes(

        uploaded_file,

        sheet_name=0,

        header=None,
    )


    if len(
        raw
    ) < 4:

        raise ValueError(
            "Template 2 資料列不足。"
        )


    # --------------------------------------------------------
    # Detect header
    # --------------------------------------------------------
    header_row = None


    for i in range(
        min(
            15,
            len(raw)
        )
    ):

        values = [

            norm_text(x)

            for x in (
                raw
                .iloc[i]
                .tolist()
            )
        ]


        if (
            "商品代號"
            in values

            and

            "商品名稱"
            in values
        ):

            header_row = i

            break


    if header_row is None:

        raise ValueError(

            "Template 2 找不到"
            "「商品代號 / 商品名稱」"
            "表頭。"
        )


    date_row = max(
        0,
        header_row - 1
    )


    headers = [

        norm_text(x)

        for x in (
            raw
            .iloc[
                header_row
            ]
            .tolist()
        )
    ]


    dates = list(

        raw
        .iloc[
            date_row
        ]
        .tolist()
    )


    # --------------------------------------------------------
    # Forward fill dates
    # --------------------------------------------------------
    ff_dates = []

    current = None


    for x in dates:

        if (
            pd.notna(x)
            and norm_text(x)
        ):

            current = x


        ff_dates.append(
            current
        )


    # --------------------------------------------------------
    # Product code index
    # --------------------------------------------------------
    product_code_idx = next(

        (
            i

            for i, h
            in enumerate(
                headers
            )

            if h
            == "商品代號"
        ),

        None,
    )


    # --------------------------------------------------------
    # Product name index
    # --------------------------------------------------------
    product_name_idx = next(

        (
            i

            for i, h
            in enumerate(
                headers
            )

            if h
            == "商品名稱"
        ),

        None,
    )


    if product_name_idx is None:

        raise ValueError(

            "Template 2"
            " 找不到商品名稱。"
        )


    target_year = (
        meta[
            "YEAR"
        ]
    )

    target_month = (
        meta[
            "MONTH"
        ]
    )


    actual_sales_idx = None


    # --------------------------------------------------------
    # Find target month's 實銷
    # --------------------------------------------------------
    for i, header in enumerate(
        headers
    ):

        if (
            "實銷"
            not in header
        ):

            continue


        date_value = (
            pd.to_datetime(
                ff_dates[i],
                errors="coerce",
            )
        )


        if (
            pd.notna(
                date_value
            )

            and

            target_year
            is not None

            and

            target_month
            is not None

            and

            date_value.year
            == target_year

            and

            date_value.month
            == target_month
        ):

            actual_sales_idx = i

            break


    # --------------------------------------------------------
    # Fallback
    # --------------------------------------------------------
    if actual_sales_idx is None:

        actual_sales_idx = next(

            (
                i

                for i, h
                in enumerate(
                    headers
                )

                if "實銷"
                in h
            ),

            None,
        )


    if actual_sales_idx is None:

        raise ValueError(

            "Template 2 找不到"
            "「實銷」欄位。"
        )


    # --------------------------------------------------------
    # Actual data
    # --------------------------------------------------------
    data = (

        raw
        .iloc[
            header_row + 1:
        ]
        .copy()
        .reset_index(
            drop=True
        )
    )


    data = (

        data
        .dropna(
            how="all"
        )
        .reset_index(
            drop=True
        )
    )


    raw_sku = (
        data
        .iloc[
            :,
            product_name_idx
        ]
    )


    qty = numeric(

        data
        .iloc[
            :,
            actual_sales_idx
        ]
    )


    valid = (

        raw_sku
        .map(
            norm_text
        )
        .ne("")
    )


    data = (

        data
        .loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    raw_sku = (

        raw_sku
        .loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    qty = (

        qty
        .loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    output = base_records(

        data,
        meta,
        "T2_MONTHLY_ACTUAL",
        uploaded_file.name,
    )


    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = ""


    output[
        "RAW_SKU"
    ] = raw_sku.map(
        norm_text
    )


    output[
        "QTY"
    ] = qty


    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    if (
        product_code_idx
        is not None
    ):

        output[
            "RAW_PRODUCT_CODE"
        ] = (

            data
            .iloc[
                :,
                product_code_idx
            ]
            .values
        )

    else:

        output[
            "RAW_PRODUCT_CODE"
        ] = ""


    return output


# ============================================================
# TEMPLATE 3A
#
# Multi-sheet transaction detail
# Example: 國泰
# ============================================================
def transform_template_3_multisheet(
    uploaded_file
):

    meta = parse_filename(
        uploaded_file.name
    )


    uploaded_file.seek(0)


    excel_file = (
        pd.ExcelFile(
            uploaded_file
        )
    )


    frames = []


    for sheet in (
        excel_file
        .sheet_names
    ):

        uploaded_file.seek(0)


        df = pd.read_excel(

            uploaded_file,

            sheet_name=sheet,

            header=0,
        )


        df = (

            df
            .dropna(
                how="all"
            )
            .reset_index(
                drop=True
            )
        )


        sku_col = (
            first_matching_column(
                df,
                [
                    "品名規格"
                ]
            )
        )


        qty_col = (
            first_matching_column(
                df,
                [
                    "數量2",
                    "數量",
                ]
            )
        )


        customer_col = (
            first_matching_column(
                df,
                [
                    "名稱"
                ]
            )
        )


        if (
            sku_col is None
            or qty_col is None
        ):

            continue


        temp = pd.DataFrame()


        temp[
            "RAW_SKU"
        ] = df[
            sku_col
        ].map(
            norm_text
        )


        temp[
            "QTY"
        ] = numeric(
            df[
                qty_col
            ]
        )


        temp[
            "RAW_ROW_CUSTOMER"
        ] = (

            df[
                customer_col
            ].map(
                norm_text
            )

            if customer_col
            is not None

            else ""
        )


        temp[
            "SOURCE_SHEET"
        ] = sheet


        temp[
            "SOURCE_INNER_ROW"
        ] = range(
            2,
            len(df) + 2
        )


        frames.append(
            temp
        )


    if not frames:

        raise ValueError(

            "Template 3 多工作表"
            "找不到"
            "「品名規格 / 數量2」。"
        )


    data = pd.concat(

        frames,

        ignore_index=True,
    )


    data = (

        data[
            data[
                "RAW_SKU"
            ].ne("")
        ]
        .reset_index(
            drop=True
        )
    )


    output = base_records(

        data,
        meta,
        "T3_TRANSACTION_MULTISHEET",
        uploaded_file.name,
    )


    output[
        "RAW_CUSTOMER"
    ] = meta[
        "FILE_RAWDATA_NAME"
    ]


    output[
        "RAW_ROW_CUSTOMER"
    ] = data[
        "RAW_ROW_CUSTOMER"
    ]


    output[
        "RAW_SKU"
    ] = data[
        "RAW_SKU"
    ]


    output[
        "QTY"
    ] = data[
        "QTY"
    ]


    output[
        "UNIT_PRICE"
    ] = pd.NA


    output[
        "TOTAL_PRICE"
    ] = pd.NA


    output[
        "SOURCE_SHEET"
    ] = data[
        "SOURCE_SHEET"
    ]


    output[
        "SOURCE_INNER_ROW"
    ] = data[
        "SOURCE_INNER_ROW"
    ]


    return output


# ============================================================
# TEMPLATE 3B
#
# Single-sheet sales detail
# Example: 加州洋酒
# ============================================================
def transform_template_3_salesdetail(
    uploaded_file
):

    meta = parse_filename(
        uploaded_file.name
    )


    raw = read_excel_bytes(

        uploaded_file,

        sheet_name=0,

        header=None,
    )


    # --------------------------------------------------------
    # Detect header row
    # --------------------------------------------------------
    header_row = None


    for i in range(
        min(
            30,
            len(raw)
        )
    ):

        values = [

            norm_text(x)

            for x in (
                raw
                .iloc[i]
                .tolist()
            )
        ]


        has_product = (
            "貨品名稱"
            in values
        )


        has_qty = any(

            "數"
            in x

            and

            "量"
            in x

            for x
            in values
        )


        if (
            has_product
            and has_qty
        ):

            header_row = i

            break


    if header_row is None:

        raise ValueError(

            "Template 3 銷貨明細"
            "找不到表頭。"
        )


    uploaded_file.seek(0)


    df = pd.read_excel(

        uploaded_file,

        sheet_name=0,

        header=header_row,
    )


    df.columns = [

        norm_text(c)

        for c
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


    # --------------------------------------------------------
    # Detect columns
    # --------------------------------------------------------
    sku_col = (
        first_matching_column(
            df,
            [
                "貨品名稱"
            ]
        )
    )


    qty_col = (
        first_matching_column(
            df,
            [
                "數 量",
                "數量",
            ]
        )
    )


    unit_price_col = (
        first_matching_column(
            df,
            [
                "單 價",
                "單價",
            ]
        )
    )


    total_price_col = (
        first_matching_column(
            df,
            [
                "總 價",
                "總價",
            ]
        )
    )


    customer_col = (
        first_matching_column(
            df,
            [
                "客戶名稱"
            ]
        )
    )


    if (
        sku_col is None
        or qty_col is None
    ):

        raise ValueError(

            "Template 3 銷貨明細"
            "找不到貨品名稱或數量。"
        )


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

        df
        .loc[
            valid
        ]
        .reset_index(
            drop=True
        )
    )


    output = base_records(

        df,
        meta,
        "T3_SALES_DETAIL",
        uploaded_file.name,
    )


    output[
        "SOURCE_ROW"
    ] = range(

        header_row + 2,

        header_row
        + 2
        + len(df)
    )


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
        ].map(
            norm_text
        )

        if customer_col
        else ""
    )


    output[
        "RAW_SKU"
    ] = df[
        sku_col
    ].map(
        norm_text
    )


    output[
        "QTY"
    ] = numeric(
        df[
            qty_col
        ]
    )


    output[
        "UNIT_PRICE"
    ] = (

        numeric(
            df[
                unit_price_col
            ]
        )

        if unit_price_col
        else pd.NA
    )


    output[
        "TOTAL_PRICE"
    ] = (

        numeric(
            df[
                total_price_col
            ]
        )

        if total_price_col
        else pd.NA
    )


    return output


# ============================================================
# Template 3 Auto Detect
# ============================================================
def detect_and_transform_template_3(
    uploaded_file
):

    uploaded_file.seek(0)


    excel_file = (
        pd.ExcelFile(
            uploaded_file
        )
    )


    # --------------------------------------------------------
    # Multi-sheet 國泰 style
    # --------------------------------------------------------
    if (
        len(
            excel_file.sheet_names
        ) > 1
    ):

        uploaded_file.seek(0)


        first_sheet = (
            pd.read_excel(

                uploaded_file,

                sheet_name=(
                    excel_file
                    .sheet_names[0]
                ),

                header=0,
            )
        )


        columns = {

            norm_col(c)

            for c
            in first_sheet.columns
        }


        if any(

            "品名規格"
            in c

            for c
            in columns
        ):

            uploaded_file.seek(0)

            return (
                transform_template_3_multisheet(
                    uploaded_file
                )
            )


    # --------------------------------------------------------
    # Single-sheet 加州 style
    # --------------------------------------------------------
    uploaded_file.seek(0)


    return (
        transform_template_3_salesdetail(
            uploaded_file
        )
    )


# ============================================================
# Mapping Helpers
# ============================================================
def resolve_required(
    df,
    required
):

    lookup = {

        norm_key(c): c

        for c
        in df.columns
    }


    resolved = {}
    missing = []


    for required_name in (
        required
    ):

        key = norm_key(
            required_name
        )


        if (
            key
            in lookup
        ):

            resolved[
                required_name
            ] = lookup[
                key
            ]

        else:

            missing.append(
                required_name
            )


    return (
        resolved,
        missing,
    )


# ============================================================
# Customer Mapping Preparation
# ============================================================
def prepare_customer_mapping(
    df
):

    resolved, missing = (
        resolve_required(
            df,
            CUSTOMER_REQUIRED,
        )
    )


    if missing:

        raise ValueError(

            "Customer Mapping 缺少："
            + ", ".join(
                missing
            )
        )


    mapping = pd.DataFrame({

        column:
            df[source]

        for column, source
        in resolved.items()
    })


    mapping[
        "OUTLET_KEY"
    ] = mapping[
        "Outlet No"
    ].map(
        norm_key
    )


    mapping[
        "RAWDATA_KEY"
    ] = mapping[
        "Rawdata Name"
    ].map(
        norm_key
    )


    mapping[
        "SALES_KEY"
    ] = mapping[
        "Sales ID"
    ].map(
        norm_key
    )


    mapping[
        "CONTRACT_KEY"
    ] = mapping[
        "Contract JDE"
    ].map(
        norm_key
    )


    duplicate_outlet = mapping[

        mapping[
            "OUTLET_KEY"
        ].ne("")

        &

        mapping.duplicated(
            "OUTLET_KEY",
            keep=False,
        )
    ].copy()


    return (
        mapping,
        duplicate_outlet,
    )


# ============================================================
# SKU Mapping Preparation
# ============================================================
def prepare_sku_mapping(
    df
):

    resolved, missing = (
        resolve_required(
            df,
            SKU_REQUIRED,
        )
    )


    if missing:

        raise ValueError(

            "SKU Mapping 缺少："
            + ", ".join(
                missing
            )
        )


    mapping = pd.DataFrame({

        column:
            df[source]

        for column, source
        in resolved.items()
    })


    mapping[
        "SKU_KEY"
    ] = mapping[
        "Mapping Name"
    ].map(
        norm_key
    )


    duplicate_mapping = mapping[

        mapping[
            "SKU_KEY"
        ].ne("")

        &

        mapping.duplicated(
            "SKU_KEY",
            keep=False,
        )
    ].copy()


    mapping_for_join = (

        mapping
        .drop_duplicates(
            "SKU_KEY",
            keep="first",
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
    preferred_sheet=None
):

    suffix = Path(
        uploaded_file.name
    ).suffix.lower()


    if (
        suffix
        == ".csv"
    ):

        return (
            read_csv_flexible(
                uploaded_file
            )
        )


    uploaded_file.seek(0)


    excel_file = (
        pd.ExcelFile(
            uploaded_file
        )
    )


    if (
        preferred_sheet
        in excel_file.sheet_names
    ):

        sheet = (
            preferred_sheet
        )

    else:

        sheet = (
            excel_file
            .sheet_names[0]
        )


    uploaded_file.seek(0)


    return pd.read_excel(

        uploaded_file,

        sheet_name=sheet,

        header=0,
    )


# ============================================================
# Attach Customer Mapping
#
# Priority:
# 1. Outlet No
# 2. Rawdata Name fallback
#
# Always LEFT JOIN
# ============================================================
def attach_customer_mapping(
    raw,
    customer_map
):

    source = raw.copy()


    source[
        "FILE_OUTLET_KEY"
    ] = source[
        "FILE_OUTLET_NO"
    ].map(
        norm_key
    )


    source[
        "FILE_RAWDATA_KEY"
    ] = source[
        "FILE_RAWDATA_NAME"
    ].map(
        norm_key
    )


    # --------------------------------------------------------
    # Match by Outlet
    # --------------------------------------------------------
    by_outlet = (

        customer_map
        .drop_duplicates(
            "OUTLET_KEY",
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


    # --------------------------------------------------------
    # Find unmatched rows
    # --------------------------------------------------------
    unmatched = (

        merged[
            "Outlet No"
        ].isna()

        |

        merged[
            "Outlet No"
        ].map(
            norm_text
        ).eq("")
    )


    # --------------------------------------------------------
    # Fallback: Rawdata Name
    # --------------------------------------------------------
    if unmatched.any():

        by_raw = (

            customer_map[

                customer_map[
                    "RAWDATA_KEY"
                ].ne("")
            ]

            .drop_duplicates(
                "RAWDATA_KEY",
                keep="first",
            )
        )


        fallback_source = (
            source.loc[
                unmatched.values
            ]
        )


        fallback = (
            fallback_source.merge(

                by_raw,

                how="left",

                left_on=(
                    "FILE_RAWDATA_KEY"
                ),

                right_on=(
                    "RAWDATA_KEY"
                ),

                suffixes=(
                    "",
                    "_CM",
                ),
            )
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


        for column in (
            mapping_columns
        ):

            if (
                column
                in fallback.columns

                and

                column
                in merged.columns
            ):

                merged.loc[
                    unmatched,
                    column
                ] = fallback[
                    column
                ].values


    # --------------------------------------------------------
    # Mapping Status
    # --------------------------------------------------------
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


    # --------------------------------------------------------
    # Sales ID validation
    # --------------------------------------------------------
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
                    norm_key(
                        row[
                            "FILE_SALES_ID"
                        ]
                    )
                    ==
                    norm_key(
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


    # --------------------------------------------------------
    # Contract validation
    # --------------------------------------------------------
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
                    norm_key(
                        row[
                            "FILE_CONTRACT_JDE"
                        ]
                    )
                    ==
                    norm_key(
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


    return merged


# ============================================================
# Attach SKU Mapping
# ============================================================
def attach_sku_mapping(
    df,
    sku_map
):

    output = df.copy()


    output[
        "SKU_KEY"
    ] = output[
        "RAW_SKU"
    ].map(
        norm_key
    )


    output = output.merge(

        sku_map,

        how="left",

        on="SKU_KEY",

        suffixes=(
            "",
            "_SKU",
        ),
    )


    output[
        "SKU_MAPPING_STATUS"
    ] = output[
        "Mapping Name"
    ].apply(

        lambda x:

        "MAPPED"

        if norm_text(x)

        else "UNMAPPED"
    )


    return output


# ============================================================
# Build Final Report
# ============================================================
def build_report(
    detail
):

    report = pd.DataFrame(
        index=detail.index
    )


    report[
        "地區"
    ] = detail[
        "REGION"
    ]


    report[
        "業務員"
    ] = detail[
        "Sales"
    ]


    report[
        "合約等級"
    ] = detail[
        "CONTRACT TYPE"
    ]


    report[
        "合約編號"
    ] = detail[
        "Contract JDE"
    ]


    report[
        "合約名稱"
    ] = detail[
        "Contract NAME"
    ]


    report[
        "店家編號"
    ] = detail[
        "Outlet No"
    ]


    report[
        "店家名稱"
    ] = detail[
        "Outlet NAME"
    ]


    report[
        "銷量(瓶)"
    ] = detail[
        "QTY"
    ]


    report[
        "單價"
    ] = detail[
        "UNIT_PRICE"
    ]


    # --------------------------------------------------------
    # Total price
    # --------------------------------------------------------
    calculated_total = (

        pd.to_numeric(
            detail[
                "QTY"
            ],
            errors="coerce",
        )

        *

        pd.to_numeric(
            detail[
                "UNIT_PRICE"
            ],
            errors="coerce",
        )
    )


    total_price = pd.to_numeric(

        detail[
            "TOTAL_PRICE"
        ],

        errors="coerce",
    )


    report[
        "總價"
    ] = total_price.where(

        total_price.notna(),

        calculated_total,
    )


    report[
        "建議售價"
    ] = pd.to_numeric(

        detail[
            "RSP"
        ],

        errors="coerce",
    )


    report[
        "價格帶"
    ] = detail[
        "Price Band"
    ]


    report[
        "製造商"
    ] = detail[
        "Manufacture"
    ]


    report[
        "品牌"
    ] = detail[
        "Band"
    ]


    report[
        "統一品項名稱"
    ] = detail[
        "Mapping Name"
    ]


    report[
        "品項"
    ] = detail[
        "RAW_SKU"
    ]


    report[
        "容量"
    ] = detail[
        "SIZE"
    ]


    report[
        "品類"
    ] = detail[
        "CATEGORY"
    ]


    report[
        "年"
    ] = pd.to_numeric(

        detail[
            "YEAR"
        ],

        errors="coerce",
    ).astype(
        "Int64"
    )


    report[
        "月"
    ] = pd.to_numeric(

        detail[
            "MONTH"
        ],

        errors="coerce",
    ).astype(
        "Int64"
    )


    return report[
        OUTPUT_COLUMNS
    ]


# ============================================================
# Export Excel
# ============================================================
def excel_bytes(
    sheets
):

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

                sheet_name=(
                    sheet_name[:31]
                ),

                index=False,
            )


    buffer.seek(0)


    return buffer.getvalue()


# ============================================================
# UI
# ============================================================
st.subheader(
    "1️⃣ 當月份資料"
)


# ============================================================
# Three Template Uploaders
# ============================================================
col1, col2, col3 = (
    st.columns(3)
)


# ------------------------------------------------------------
# Template 1
# ------------------------------------------------------------
with col1:

    st.markdown(
        "### Template 1｜標準 POS SKU 表"
    )

    st.caption(
        "例如：冠德系列。支援 CSV / Excel。"
    )


    template_1_files = (
        st.file_uploader(

            "上傳 Template 1",

            type=[
                "csv",
                "xlsx",
                "xls",
            ],

            accept_multiple_files=True,

            key="template_1",
        )
    )


# ------------------------------------------------------------
# Template 2
# ------------------------------------------------------------
with col2:

    st.markdown(
        "### Template 2｜月份實銷表"
    )

    st.caption(
        "例如：洋酒城系列。"
        "會依照檔名年月抓對應月份的實銷。"
    )


    template_2_files = (
        st.file_uploader(

            "上傳 Template 2",

            type=[
                "xlsx",
                "xls",
            ],

            accept_multiple_files=True,

            key="template_2",
        )
    )


# ------------------------------------------------------------
# Template 3
# ------------------------------------------------------------
with col3:

    st.markdown(
        "### Template 3｜銷售明細表"
    )

    st.caption(
        "例如：國泰 / 加州洋酒。"
        "自動辨識多 Sheet 或單 Sheet 格式。"
    )


    template_3_files = (
        st.file_uploader(

            "上傳 Template 3",

            type=[
                "xlsx",
                "xls",
            ],

            accept_multiple_files=True,

            key="template_3",
        )
    )


# ============================================================
# Mapping Files
# ============================================================
st.divider()


st.subheader(
    "2️⃣ Mapping 資料"
)


map_col1, map_col2 = (
    st.columns(2)
)


with map_col1:

    customer_file = (
        st.file_uploader(

            "Customer Mapping",

            type=[
                "xlsx",
                "xls",
                "csv",
            ],

            key="customer_mapping",
        )
    )


with map_col2:

    sku_file = (
        st.file_uploader(

            "SKU Mapping",

            type=[
                "xlsx",
                "xls",
                "csv",
            ],

            key="sku_mapping",
        )
    )


st.caption(
    "如果 Mapping 放在同一個 Excel，"
    "Customer Mapping 會優先找 Customer data sheet，"
    "SKU Mapping 會優先找 SKU data sheet。"
)


# ============================================================
# Validation
# ============================================================
all_template_files = (

    (template_1_files or [])

    +

    (template_2_files or [])

    +

    (template_3_files or [])
)


if not all_template_files:

    st.info(
        "請至少上傳一個 Template。"
    )

    st.stop()


if (
    customer_file is None
    or sku_file is None
):

    st.info(
        "請上傳 Customer Mapping "
        "與 SKU Mapping。"
    )

    st.stop()


# ============================================================
# Load Customer Mapping
# ============================================================
try:

    customer_raw = (
        read_mapping_file(

            customer_file,

            preferred_sheet=(
                "Customer data"
            ),
        )
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

    sku_raw = (
        read_mapping_file(

            sku_file,

            preferred_sheet=(
                "SKU data"
            ),
        )
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
# Process Templates
# ============================================================
frames = []

upload_logs = []


def run_group(
    files,
    label,
    transformer,
):


    for file in (
        files or []
    ):


        meta = parse_filename(
            file.name
        )


        try:

            file.seek(0)


            data = transformer(
                file
            )


            frames.append(
                data
            )


            upload_logs.append({

                "Template":
                    label,

                "檔案":
                    file.name,

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

                "Template":
                    label,

                "檔案":
                    file.name,

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


# ============================================================
# Run 3 Template Groups
# ============================================================
run_group(

    template_1_files,

    "Template 1",

    transform_template_1,
)


run_group(

    template_2_files,

    "Template 2",

    transform_template_2,
)


run_group(

    template_3_files,

    "Template 3",

    detect_and_transform_template_3,
)


upload_log = pd.DataFrame(
    upload_logs
)


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
# Combine Source Data
# ============================================================
raw_all = pd.concat(

    frames,

    ignore_index=True,

    sort=False,
)


# ============================================================
# Customer Mapping
# ============================================================
with_customer = (
    attach_customer_mapping(

        raw_all,

        customer_map,
    )
)


# ============================================================
# SKU Mapping
# ============================================================
detail = (
    attach_sku_mapping(

        with_customer,

        sku_map,
    )
)


# ============================================================
# Build Final Report
# ============================================================
final_report = (
    build_report(
        detail
    )
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

    .eq(
        "UNMAPPED"
    )

    .sum()
)


sku_unmapped = int(

    detail[
        "SKU_MAPPING_STATUS"
    ]

    .eq(
        "UNMAPPED"
    )

    .sum()
)


sales_mismatch = int(

    detail[
        "SALES_ID_CHECK"
    ]

    .eq(
        "MISMATCH"
    )

    .sum()
)


contract_mismatch = int(

    detail[
        "CONTRACT_CHECK"
    ]

    .eq(
        "MISMATCH"
    )

    .sum()
)


failed_files = int(

    upload_log[
        "狀態"
    ]

    .eq(
        "FAILED"
    )

    .sum()
)


# ============================================================
# Dashboard
# ============================================================
st.divider()


st.subheader(
    "📊 Dashboard"
)


d1, d2, d3, d4, d5, d6 = (
    st.columns(6)
)


d1.metric(
    "來源資料",
    f"{source_count:,}"
)


d2.metric(
    "輸出資料",
    f"{output_count:,}",
    delta=(
        f"{output_count-source_count:+,}"
    ),
)


d3.metric(
    "Customer Unmapped",
    f"{customer_unmapped:,}"
)


d4.metric(
    "SKU Unmapped",
    f"{sku_unmapped:,}"
)


d5.metric(
    "上傳失敗",
    f"{failed_files:,}"
)


d6.metric(
    "檔案數",
    f"{len(upload_log):,}"
)


# ============================================================
# Reconciliation Check
# ============================================================
if (
    source_count
    == output_count
):

    st.success(

        f"✅ 資料筆數一致："
        f"來源 {source_count:,} "
        f"= 輸出 {output_count:,}。"

        " Unmapped 資料仍保留，"
        "不會因 Mapping 失敗被刪除。"
    )


else:

    st.error(

        f"❌ 筆數異常："
        f"來源 {source_count:,}，"
        f"輸出 {output_count:,}。"
    )


# ============================================================
# Validation Metrics
# ============================================================
check_col1, check_col2 = (
    st.columns(2)
)


check_col1.metric(

    "業務代號不一致",

    f"{sales_mismatch:,}"
)


check_col2.metric(

    "合約編號不一致",

    f"{contract_mismatch:,}"
)


# ============================================================
# Submission Tracking
# ============================================================
st.subheader(
    "📥 業務繳交狀況"
)


submission = (

    upload_log

    .groupby(

        [
            "業務代號",
            "合約編號",
            "年月",
        ],

        dropna=False,

        as_index=False,
    )

    .agg(

        檔案數=(
            "檔案",
            "count"
        ),

        成功檔案=(
            "狀態",
            lambda x:
                int(
                    (
                        x
                        == "SUCCESS"
                    ).sum()
                )
        ),

        失敗檔案=(
            "狀態",
            lambda x:
                int(
                    (
                        x
                        == "FAILED"
                    ).sum()
                )
        ),

        資料筆數=(
            "原始筆數",
            "sum"
        ),
    )
)


submission[
    "繳交狀態"
] = submission.apply(

    lambda row:

    (
        "成功"

        if (
            row[
                "失敗檔案"
            ] == 0

            and

            row[
                "成功檔案"
            ] > 0
        )

        else (

            "部分失敗"

            if (
                row[
                    "成功檔案"
                ] > 0
            )

            else "失敗"
        )
    ),

    axis=1,
)


st.dataframe(

    submission,

    use_container_width=True,

    hide_index=True,
)


# ============================================================
# Upload Status
# ============================================================
st.subheader(
    "📁 上傳狀況"
)


st.dataframe(

    upload_log,

    use_container_width=True,

    hide_index=True,
)


# ============================================================
# Preview
# ============================================================
st.subheader(
    "👀 Preview"
)


audit_columns = [

    "SOURCE_FILE",
    "SOURCE_ROW",
    "TEMPLATE_TYPE",

    "FILE_SALES_ID",
    "FILE_CONTRACT_JDE",
    "FILE_OUTLET_NO",
    "FILE_RAWDATA_NAME",

    "RAW_CUSTOMER",
    "RAW_ROW_CUSTOMER",
    "RAW_SKU",

    "CUSTOMER_MAPPING_STATUS",
    "SKU_MAPPING_STATUS",

    "SALES_ID_CHECK",
    "CONTRACT_CHECK",
]


audit_columns = [

    column

    for column
    in audit_columns

    if column
    in detail.columns
]


tabs = st.tabs([

    "最終報表",

    "Customer Unmapped",

    "SKU Unmapped",

    "上傳明細",

    "Customer Mapping 重複",

    "SKU Mapping 重複",
])


# ============================================================
# Final Report Preview
# ============================================================
with tabs[0]:

    st.dataframe(

        final_report,

        use_container_width=True,

        hide_index=True,

        height=550,
    )


# ============================================================
# Customer Unmapped Preview
# ============================================================
with tabs[1]:

    mask = (

        detail[
            "CUSTOMER_MAPPING_STATUS"
        ]

        .eq(
            "UNMAPPED"
        )
    )


    preview = pd.concat(

        [

            detail.loc[
                mask,
                audit_columns
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


# ============================================================
# SKU Unmapped Preview
# ============================================================
with tabs[2]:

    mask = (

        detail[
            "SKU_MAPPING_STATUS"
        ]

        .eq(
            "UNMAPPED"
        )
    )


    preview = pd.concat(

        [

            detail.loc[
                mask,
                audit_columns
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


# ============================================================
# Upload Detail
# ============================================================
with tabs[3]:

    st.dataframe(

        upload_log,

        use_container_width=True,

        hide_index=True,

        height=520,
    )


# ============================================================
# Customer Mapping Duplicate
# ============================================================
with tabs[4]:

    st.write(
        f"共 {len(customer_duplicates):,} 筆"
    )


    st.dataframe(

        customer_duplicates,

        use_container_width=True,

        hide_index=True,

        height=520,
    )


# ============================================================
# SKU Mapping Duplicate
# ============================================================
with tabs[5]:

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
# Export
# ============================================================
st.divider()


st.subheader(
    "⬇️ Export"
)


unmapped_mask = (

    detail[
        "CUSTOMER_MAPPING_STATUS"
    ].eq(
        "UNMAPPED"
    )

    |

    detail[
        "SKU_MAPPING_STATUS"
    ].eq(
        "UNMAPPED"
    )
)


unmapped = pd.concat(

    [

        detail.loc[
            unmapped_mask,
            audit_columns
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
# Excel Workbook
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


export_col1, export_col2, export_col3 = (
    st.columns(3)
)


# ============================================================
# Download Full Excel
# ============================================================
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


# ============================================================
# Download Final CSV
# ============================================================
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


# ============================================================
# Download Unmapped Excel
# ============================================================
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
