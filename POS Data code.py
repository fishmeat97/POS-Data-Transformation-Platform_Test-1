import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# Page Config
# ============================================================
st.set_page_config(
    page_title="POS Monthly Data Mapping",
    page_icon="📊",
    layout="wide",
)

st.title("📊 POS Monthly Data Mapping")
st.caption(
    "單一入口上傳多個月度 Template，系統自動辨識格式，"
    "再執行 Customer Mapping、SKU Mapping、資料勾稽與報表輸出。"
)


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
    "Manufacture",
    "Band",
    "RSP",
    "Price Band",
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
def transform_template_2(uploaded_file):
    meta = parse_filename(
        uploaded_file.name
    )

    raw = read_excel_bytes(
        uploaded_file,
        sheet_name=0,
        header=None,
    )

    if len(raw) < 4:
        raise ValueError(
            "Template 2 資料列不足。"
        )

    header_row = None

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
        ):
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            "Template 2 找不到「商品代號 / 商品名稱」表頭。"
        )

    date_row = max(
        0,
        header_row - 1,
    )

    headers = [
        norm_text(x)
        for x in raw.iloc[
            header_row
        ].tolist()
    ]

    dates = list(
        raw.iloc[
            date_row
        ].tolist()
    )

    # forward fill date headers
    ff_dates = []
    current = None

    for x in dates:
        if (
            pd.notna(x)
            and norm_text(x)
        ):
            current = x

        ff_dates.append(current)

    product_code_idx = next(
        (
            i
            for i, h in enumerate(headers)
            if h == "商品代號"
        ),
        None,
    )

    product_name_idx = next(
        (
            i
            for i, h in enumerate(headers)
            if h == "商品名稱"
        ),
        None,
    )

    if product_name_idx is None:
        raise ValueError(
            "Template 2 找不到商品名稱。"
        )

    target_year = meta["YEAR"]
    target_month = meta["MONTH"]

    actual_sales_idx = None

    for i, header in enumerate(headers):
        if "實銷" not in header:
            continue

        date_value = pd.to_datetime(
            ff_dates[i],
            errors="coerce",
        )

        if (
            pd.notna(date_value)
            and target_year is not None
            and target_month is not None
            and date_value.year == target_year
            and date_value.month == target_month
        ):
            actual_sales_idx = i
            break

    # fallback to first actual-sales column
    if actual_sales_idx is None:
        actual_sales_idx = next(
            (
                i
                for i, h in enumerate(headers)
                if "實銷" in h
            ),
            None,
        )

    if actual_sales_idx is None:
        raise ValueError(
            "Template 2 找不到「實銷」欄位。"
        )

    data = (
        raw
        .iloc[
            header_row + 1:
        ]
        .copy()
        .reset_index(drop=True)
    )

    data = (
        data
        .dropna(how="all")
        .reset_index(drop=True)
    )

    raw_sku = data.iloc[
        :,
        product_name_idx
    ]

    qty = numeric(
        data.iloc[
            :,
            actual_sales_idx
        ]
    )

    valid = (
        raw_sku
        .map(norm_text)
        .ne("")
    )

    data = (
        data
        .loc[valid]
        .reset_index(drop=True)
    )

    raw_sku = (
        raw_sku
        .loc[valid]
        .reset_index(drop=True)
    )

    qty = (
        qty
        .loc[valid]
        .reset_index(drop=True)
    )

    output = base_records(
        data,
        meta,
        "T2_MONTHLY_ACTUAL",
        uploaded_file.name,
    )

    output["RAW_CUSTOMER"] = meta[
        "FILE_RAWDATA_NAME"
    ]

    output["RAW_ROW_CUSTOMER"] = ""

    output["RAW_SKU"] = (
        raw_sku
        .map(norm_text)
    )

    output["QTY"] = qty
    output["UNIT_PRICE"] = pd.NA
    output["TOTAL_PRICE"] = pd.NA

    if product_code_idx is not None:
        output["RAW_PRODUCT_CODE"] = (
            data.iloc[
                :,
                product_code_idx
            ].values
        )
    else:
        output["RAW_PRODUCT_CODE"] = ""

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
def prepare_customer_mapping(df):
    resolved, missing = resolve_required(
        df,
        CUSTOMER_REQUIRED,
    )

    if missing:
        raise ValueError(
            "Customer Mapping 缺少："
            + ", ".join(missing)
        )

    mapping = pd.DataFrame({
        column: df[source]
        for column, source in resolved.items()
    })

    mapping["OUTLET_KEY"] = (
        mapping["Outlet No"]
        .map(norm_key)
    )

    mapping["RAWDATA_KEY"] = (
        mapping["Rawdata Name"]
        .map(norm_key)
    )

    mapping["SALES_KEY"] = (
        mapping["Sales ID"]
        .map(norm_key)
    )

    mapping["CONTRACT_KEY"] = (
        mapping["Contract JDE"]
        .map(norm_key)
    )

    duplicate_outlet = mapping[
        mapping["OUTLET_KEY"].ne("")
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
# SKU Mapping
# ============================================================
def prepare_sku_mapping(df):
    resolved, missing = resolve_required(
        df,
        SKU_REQUIRED,
    )

    if missing:
        raise ValueError(
            "SKU Mapping 缺少："
            + ", ".join(missing)
        )

    mapping = pd.DataFrame({
        column: df[source]
        for column, source in resolved.items()
    })

    mapping["SKU_KEY"] = (
        mapping["Mapping Name"]
        .map(norm_key)
    )

    duplicate_mapping = mapping[
        mapping["SKU_KEY"].ne("")
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
def attach_customer_mapping(
    raw,
    customer_map,
):
    source = raw.copy()

    source["FILE_OUTLET_KEY"] = (
        source["FILE_OUTLET_NO"]
        .map(norm_key)
    )

    source["FILE_RAWDATA_KEY"] = (
        source["FILE_RAWDATA_NAME"]
        .map(norm_key)
    )

    # First match by Outlet No
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

    unmatched = (
        merged["Outlet No"].isna()
        |
        merged["Outlet No"]
        .map(norm_text)
        .eq("")
    )

    # Fallback by Rawdata Name
    if unmatched.any():
        by_raw = (
            customer_map[
                customer_map["RAWDATA_KEY"].ne("")
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
                and column in merged.columns
            ):
                merged.loc[
                    unmatched,
                    column
                ] = fallback[
                    column
                ].values

    # Customer Mapping Status
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

    # Sales ID Check
    merged[
        "SALES_ID_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if row[
                "CUSTOMER_MAPPING_STATUS"
            ] == "UNMAPPED"
            else (
                "MATCH"
                if norm_key(
                    row["FILE_SALES_ID"]
                )
                ==
                norm_key(
                    row["Sales ID"]
                )
                else "MISMATCH"
            )
        ),
        axis=1,
    )

    # Contract Check
    merged[
        "CONTRACT_CHECK"
    ] = merged.apply(
        lambda row:
        (
            "UNMAPPED"
            if row[
                "CUSTOMER_MAPPING_STATUS"
            ] == "UNMAPPED"
            else (
                "MATCH"
                if norm_key(
                    row["FILE_CONTRACT_JDE"]
                )
                ==
                norm_key(
                    row["Contract JDE"]
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
def attach_sku_mapping(
    df,
    sku_map,
):
    output = df.copy()

    output["SKU_KEY"] = (
        output["RAW_SKU"]
        .map(norm_key)
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
def build_report(detail):
    report = pd.DataFrame(
        index=detail.index
    )

    report["地區"] = detail[
        "REGION"
    ]

    report["業務員"] = detail[
        "Sales"
    ]

    report["合約等級"] = detail[
        "CONTRACT TYPE"
    ]

    report["合約編號"] = detail[
        "Contract JDE"
    ]

    report["合約名稱"] = detail[
        "Contract NAME"
    ]

    report["店家編號"] = detail[
        "Outlet No"
    ]

    report["店家名稱"] = detail[
        "Outlet NAME"
    ]

    report["銷量(瓶)"] = detail[
        "QTY"
    ]

    report["單價"] = detail[
        "UNIT_PRICE"
    ]

    calculated_total = (
        pd.to_numeric(
            detail["QTY"],
            errors="coerce",
        )
        *
        pd.to_numeric(
            detail["UNIT_PRICE"],
            errors="coerce",
        )
    )

    total_price = pd.to_numeric(
        detail["TOTAL_PRICE"],
        errors="coerce",
    )

    report["總價"] = (
        total_price.where(
            total_price.notna(),
            calculated_total,
        )
    )

    report["建議售價"] = pd.to_numeric(
        detail["RSP"],
        errors="coerce",
    )

    report["價格帶"] = detail[
        "Price Band"
    ]

    report["製造商"] = detail[
        "Manufacture"
    ]

    report["品牌"] = detail[
        "Band"
    ]

    report["統一品項名稱"] = detail[
        "Mapping Name"
    ]

    report["品項"] = detail[
        "RAW_SKU"
    ]

    report["容量"] = detail[
        "SIZE"
    ]

    report["品類"] = detail[
        "CATEGORY"
    ]

    report["年"] = pd.to_numeric(
        detail["YEAR"],
        errors="coerce",
    ).astype("Int64")

    report["月"] = pd.to_numeric(
        detail["MONTH"],
        errors="coerce",
    ).astype("Int64")

    return report[
        OUTPUT_COLUMNS
    ]


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
st.subheader(
    "1️⃣ 上傳當月份 Template"
)

template_files = st.file_uploader(
    "可同時上傳多個 Template，系統會自動辨識格式",
    type=[
        "csv",
        "xlsx",
        "xls",
    ],
    accept_multiple_files=True,
    key="monthly_templates",
)

st.caption(
    "支援目前已確認的冠德、洋酒城、國泰、加州洋酒等格式；"
    "使用者不需要手動選擇 Template 類型。"
)


# ============================================================
# UI - Mapping Upload
# ============================================================
st.divider()

st.subheader(
    "2️⃣ 上傳 Mapping 資料"
)

map_col1, map_col2 = (
    st.columns(2)
)

with map_col1:
    customer_file = st.file_uploader(
        "Customer Mapping",
        type=[
            "xlsx",
            "xls",
            "csv",
        ],
        key="customer_mapping",
    )

with map_col2:
    sku_file = st.file_uploader(
        "SKU Mapping",
        type=[
            "xlsx",
            "xls",
            "csv",
        ],
        key="sku_mapping",
    )

st.caption(
    "若 Mapping 放在同一份 Excel，"
    "Customer Mapping 會優先找 Customer data sheet，"
    "SKU Mapping 會優先找 SKU data sheet。"
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
    f"{source_count:,}",
)

d2.metric(
    "輸出資料",
    f"{output_count:,}",
    delta=(
        f"{output_count - source_count:+,}"
    ),
)

d3.metric(
    "Customer Unmapped",
    f"{customer_unmapped:,}",
)

d4.metric(
    "SKU Unmapped",
    f"{sku_unmapped:,}",
)

d5.metric(
    "成功檔案",
    f"{successful_files:,}",
)

d6.metric(
    "失敗檔案",
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
        f"✅ 資料筆數一致：來源 {source_count:,} = 輸出 {output_count:,}。"
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
# DASHBOARD
# ============================================================

st.divider()

st.title("📊 Monthly POS Dashboard")


# ============================================================
# 1. Detect Current Reporting Period
# ============================================================

valid_periods = (
    upload_log["年月"]
    .astype(str)
    .replace("", pd.NA)
    .dropna()
    .unique()
)

if len(valid_periods) > 0:

    reporting_period = sorted(
        valid_periods
    )[-1]

else:

    reporting_period = ""


# ============================================================
# 2. Build Expected Submission List
#
# Customer Mapping = expected submission population
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


# ------------------------------------------------------------
# Normalize Keys
# ------------------------------------------------------------

expected_submission[
    "SALES_KEY"
] = (
    expected_submission[
        "Sales ID"
    ]
    .map(norm_key)
)


expected_submission[
    "CONTRACT_KEY"
] = (
    expected_submission[
        "Contract JDE"
    ]
    .map(norm_key)
)


expected_submission[
    "OUTLET_KEY"
] = (
    expected_submission[
        "Outlet No"
    ]
    .map(norm_key)
)


expected_submission[
    "EXPECTED_KEY"
] = (

    expected_submission[
        "SALES_KEY"
    ]

    + "|"

    + expected_submission[
        "CONTRACT_KEY"
    ]

    + "|"

    + expected_submission[
        "OUTLET_KEY"
    ]
)


# ------------------------------------------------------------
# Remove duplicate expected submissions
# ------------------------------------------------------------

expected_submission = (
    expected_submission
    .drop_duplicates(
        subset=[
            "EXPECTED_KEY"
        ]
    )
    .reset_index(
        drop=True
    )
)


# ============================================================
# 3. Build Actual Upload Key
# ============================================================

actual_upload = (
    upload_log
    .copy()
)


actual_upload[
    "SALES_KEY"
] = (
    actual_upload[
        "業務代號"
    ]
    .map(norm_key)
)


actual_upload[
    "CONTRACT_KEY"
] = (
    actual_upload[
        "合約編號"
    ]
    .map(norm_key)
)


actual_upload[
    "OUTLET_KEY"
] = (
    actual_upload[
        "店家編號"
    ]
    .map(norm_key)
)


actual_upload[
    "UPLOAD_KEY"
] = (

    actual_upload[
        "SALES_KEY"
    ]

    + "|"

    + actual_upload[
        "CONTRACT_KEY"
    ]

    + "|"

    + actual_upload[
        "OUTLET_KEY"
    ]
)


# ============================================================
# 4. Merge Expected vs Actual
# ============================================================

submission_detail = (
    expected_submission
    .merge(

        actual_upload[
            [
                "UPLOAD_KEY",
                "檔案",
                "辨識格式",
                "狀態",
                "錯誤",
                "原始筆數",
                "年月",
                "檔名格式",
            ]
        ],

        how="left",

        left_on="EXPECTED_KEY",

        right_on="UPLOAD_KEY",
    )
)


# ============================================================
# 5. Submission Status
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
# 6. Dashboard KPI
# ============================================================

expected_count = len(
    expected_submission
)


success_count = int(
    (
        submission_detail[
            "SUBMISSION_STATUS"
        ]
        == "SUCCESS"
    )
    .sum()
)


missing_count = int(
    (
        submission_detail[
            "SUBMISSION_STATUS"
        ]
        == "MISSING"
    )
    .sum()
)


submission_failed_count = int(
    (
        submission_detail[
            "SUBMISSION_STATUS"
        ]
        == "FAILED"
    )
    .sum()
)


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


# ============================================================
# KPI ROW 1
# ============================================================

st.subheader(
    f"📅 Reporting Period: {reporting_period}"
)


kpi1, kpi2, kpi3, kpi4 = (
    st.columns(4)
)


kpi1.metric(
    "📋 本月應繳",
    f"{expected_count:,}",
)


kpi2.metric(
    "✅ 成功繳交",
    f"{success_count:,}",
)


kpi3.metric(
    "⏳ 尚未繳交",
    f"{missing_count:,}",
)


kpi4.metric(
    "❌ 上傳失敗",
    f"{submission_failed_count:,}",
)


# ============================================================
# KPI ROW 2
# ============================================================

kpi5, kpi6, kpi7, kpi8 = (
    st.columns(4)
)


kpi5.metric(
    "📥 原始資料筆數",
    f"{source_count:,}",
)


kpi6.metric(
    "📤 最終輸出筆數",
    f"{output_count:,}",
    delta=(
        f"{output_count - source_count:+,}"
    ),
)


kpi7.metric(
    "🧩 SKU Unmapped",
    f"{sku_unmapped:,}",
)


kpi8.metric(
    "🏪 Customer Unmapped",
    f"{customer_unmapped:,}",
)


# ============================================================
# 7. Data Reconciliation
# ============================================================

st.subheader(
    "🔐 Data Reconciliation"
)


if (
    source_count
    == output_count
):

    st.success(

        f"""
        ✅ 資料筆數核對成功

        原始資料：{source_count:,} 筆  
        最終輸出：{output_count:,} 筆  

        差異：0 筆
        """
    )

else:

    difference = (
        output_count
        - source_count
    )

    st.error(

        f"""
        ❌ 資料筆數不一致

        原始資料：{source_count:,} 筆  
        最終輸出：{output_count:,} 筆  

        差異：{difference:+,} 筆
        """
    )


# ============================================================
# 8. Mapping Completion Rate
# ============================================================

st.subheader(
    "🧩 Mapping Completion"
)


if source_count > 0:

    customer_mapping_rate = (

        (
            source_count
            - customer_unmapped
        )
        / source_count
    )


    sku_mapping_rate = (

        (
            source_count
            - sku_unmapped
        )
        / source_count
    )

else:

    customer_mapping_rate = 0

    sku_mapping_rate = 0


mapping_col1, mapping_col2 = (
    st.columns(2)
)


with mapping_col1:

    st.write(
        "Customer Mapping"
    )

    st.progress(
        customer_mapping_rate
    )

    st.caption(

        f"""
        {customer_mapping_rate:.2%}
        Mapped

        Unmapped:
        {customer_unmapped:,}
        """
    )


with mapping_col2:

    st.write(
        "SKU Mapping"
    )

    st.progress(
        sku_mapping_rate
    )

    st.caption(

        f"""
        {sku_mapping_rate:.2%}
        Mapped

        Unmapped:
        {sku_unmapped:,}
        """
    )


# ============================================================
# 9. Sales Submission Progress
# ============================================================

st.subheader(
    "👤 業務繳交進度"
)


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
                        x
                        == "SUCCESS"
                    ).sum()
                ),
        ),

        失敗=(
            "SUBMISSION_STATUS",
            lambda x:
                int(
                    (
                        x
                        == "FAILED"
                    ).sum()
                ),
        ),

        未繳=(
            "SUBMISSION_STATUS",
            lambda x:
                int(
                    (
                        x
                        == "MISSING"
                    ).sum()
                ),
        ),
    )
)


sales_submission[
    "完成率"
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
    "完成率"
] = (

    sales_submission[
        "完成率"
    ]
    .fillna(0)
)


sales_submission[
    "完成率 %"
] = (

    sales_submission[
        "完成率"
    ]
    .apply(
        lambda x:
            f"{x:.1%}"
    )
)


# ============================================================
# Overall Submission Progress
# ============================================================

if expected_count > 0:

    overall_completion = (
        success_count
        / expected_count
    )

else:

    overall_completion = 0


st.write(
    "### Overall Submission Progress"
)


st.progress(
    overall_completion
)


st.caption(
    f"{success_count:,} / {expected_count:,} "
    f"({overall_completion:.1%})"
)


# ============================================================
# Sales Table
# ============================================================

st.dataframe(

    sales_submission[
        [
            "Sales ID",
            "Sales",
            "應繳",
            "成功",
            "失敗",
            "未繳",
            "完成率 %",
        ]
    ],

    use_container_width=True,

    hide_index=True,
)


# ============================================================
# 10. Submission Status Distribution
# ============================================================

st.subheader(
    "📊 繳交狀況分布"
)


status_summary = (

    submission_detail[
        "SUBMISSION_STATUS"
    ]

    .value_counts()

    .rename_axis(
        "Status"
    )

    .reset_index(
        name="Count"
    )
)


st.bar_chart(
    status_summary,
    x="Status",
    y="Count",
)


# ============================================================
# 11. Shop Submission Detail
# ============================================================

st.subheader(
    "🏪 店家繳交明細"
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
            "年月",
            "檔案",
            "辨識格式",
            "SUBMISSION_STATUS",
            "錯誤",
        ]
    ]
    .copy()
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

            "Rawdata Name":
                "Rawdata Name",

            "SUBMISSION_STATUS":
                "繳交狀態",
        }
    )
)


# ============================================================
# Status Filter
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
# 12. Upload Health
# ============================================================

st.subheader(
    "📁 檔案上傳狀況"
)


upload_health = (
    upload_log[
        [
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
    ]
)


st.dataframe(

    upload_health,

    use_container_width=True,

    hide_index=True,

    height=400,
)


# ============================================================
# Upload Failure Warning
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

        f"""
        ⚠️ 有 {len(failed_uploads):,} 個檔案處理失敗，
        請至「檔案上傳狀況」查看錯誤訊息。
        """
    )


# ============================================================
# 13. Mapping Exceptions
# ============================================================

st.subheader(
    "⚠️ Data Quality Exceptions"
)


exception1, exception2, exception3 = (
    st.columns(3)
)


exception1.metric(

    "SKU Unmapped",

    f"{sku_unmapped:,}"
)


exception2.metric(

    "Customer Unmapped",

    f"{customer_unmapped:,}"
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


exception3.metric(

    "檔名格式異常",

    f"{filename_error_count:,}"
)

# ============================================================
# Preview
# ============================================================
st.subheader(
    "👀 Preview"
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

st.subheader(
    "⬇️ Export"
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
