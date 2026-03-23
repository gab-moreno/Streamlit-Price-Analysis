import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from build_comparison import (
    PROVINCE_TAX,
    _LABELS,
    all_items_with_subitems,
    all_suppliers,
    build_individual,
    build_package_detail,
    first_brand,
    groups_from,
    is_service,
    items_for,
    price_for,
    price_for_any_type,
    uniq_col,
)

st.set_page_config(layout="wide")

REQUIRED_COLUMNS = ["type", "supplier", "brand", "code", "description", "Power Type", "price"]


def label(lang: str, key: str) -> str:
    return _LABELS.get(lang, _LABELS["en"]).get(key, _LABELS["en"].get(key, key))


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in ["type", "supplier", "brand", "code", "description", "Power Type"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    if "price" in df.columns:
        df["price"] = pd.to_numeric(df["price"], errors="coerce").fillna(0.0)
    return df


def currency(value: float | int | None) -> str:
    if value is None or value == "":
        return ""
    return f"${float(value):,.2f}"


def marked_up_price(base_price: float, supplier: str, markup: dict[str, float], description: str) -> float:
    pct = markup.get(supplier, 0.0) / 100.0
    return base_price if is_service(str(description)) else base_price * (1 + pct)


def preview_styles() -> str:
    return """
    <style>
      .pa-table {border-collapse: collapse; width: 100%; margin: 0 0 2rem 0; font-family: Arial, sans-serif;}
      .pa-table th, .pa-table td {border: 1px solid #d9d9d9; padding: 6px 8px; vertical-align: middle;}
      .pa-table th {background: #f8f9fa; color: #555; font-size: 0.9rem;}
      .winner-col, .winner-head {background: #f2faf2;}
      .details-cell {background: #fafafa; white-space: pre-line; vertical-align: top; min-width: 180px;}
      .image-cell {background: #fff; min-width: 80px;}
      .num {text-align: right; white-space: nowrap;}
      .center {text-align: center;}
      .subitem {color: #555; font-style: italic; padding-left: 1.25rem;}
      .summary-label {font-weight: 700; text-align: right;}
      .best-row td {background: #f2faf2; font-weight: 700; color: #1b5e20; text-align: center;}
    </style>
    """


def build_individual_preview(df: pd.DataFrame, cfg: dict) -> str:
    sections: list[str] = [preview_styles()]
    tr = cfg["tax_rate"] / 100.0
    tl = cfg["tax_label"]
    show_toggle = cfg["include_toggle"]
    markup = cfg["markup"]

    for code, power_type in groups_from(df):
        items = items_for(df, code, power_type)
        suppliers = uniq_col(items, "supplier")
        descriptions = uniq_col(items, "description")
        brand = first_brand(items)

        totals_before_tax: dict[str, float] = {}
        for supplier in suppliers:
            total = 0.0
            for desc in descriptions:
                base = price_for(items, supplier, desc)
                if base > 0:
                    line_total = marked_up_price(base, supplier, markup, desc)
                    if not show_toggle or True:
                        total += line_total
            totals_before_tax[supplier] = total

        winner = min(
            (s for s in suppliers if totals_before_tax.get(s, 0) > 0),
            key=lambda s: totals_before_tax[s] * (1 + tr),
            default="",
        )

        header = [label(cfg["lang"], "details"), label(cfg["lang"], "image")]
        if show_toggle:
            header.append(label(cfg["lang"], "incl_in_total"))
        header.extend([label(cfg["lang"], "qty"), label(cfg["lang"], "line_item")])
        header.extend(suppliers)

        rows = [f"<h4>{code} — {power_type}</h4>", "<table class='pa-table'>", "<thead><tr>"]
        for heading in header:
            cls = "winner-head" if heading == winner else ""
            rows.append(f"<th class='{cls}'>{heading}</th>")
        rows.append("</tr></thead><tbody>")

        details = (
            f"{label(cfg['lang'], 'brand_label')}\n{brand}\n\n"
            f"{label(cfg['lang'], 'code_label')}\n{code}\n\n"
            f"{label(cfg['lang'], 'power_label')}\n{power_type}"
        )
        span = len(descriptions) + 2

        for idx, desc in enumerate(descriptions):
            rows.append("<tr>")
            if idx == 0:
                rows.append(f"<td class='details-cell' rowspan='{span}'>{details}</td>")
                rows.append(f"<td class='image-cell' rowspan='{span}'></td>")
            if show_toggle:
                rows.append("<td class='center'>✓</td>")
            rows.append("<td class='center'>1</td>")
            rows.append(f"<td>{desc}</td>")
            for supplier in suppliers:
                value = price_for(items, supplier, desc)
                display = currency(marked_up_price(value, supplier, markup, desc)) if value > 0 else ""
                cls = "num winner-col" if supplier == winner else "num"
                rows.append(f"<td class='{cls}'>{display}</td>")
            rows.append("</tr>")

        tax_totals = {s: totals_before_tax[s] * tr for s in suppliers}
        grand_totals = {s: totals_before_tax[s] + tax_totals[s] for s in suppliers}

        for title_key, source in [
            ("total_before_tax", totals_before_tax),
            ("tax", tax_totals),
            ("total", grand_totals),
        ]:
            rows.append("<tr>")
            if show_toggle:
                rows.append(f"<td colspan='3' class='summary-label'>{label(cfg['lang'], title_key)} {tl if title_key == 'tax' else ''}</td>")
            else:
                rows.append(f"<td colspan='2' class='summary-label'>{label(cfg['lang'], title_key)} {tl if title_key == 'tax' else ''}</td>")
            for supplier in suppliers:
                cls = "num winner-col" if supplier == winner else "num"
                rows.append(f"<td class='{cls}'>{currency(source[supplier])}</td>")
            rows.append("</tr>")

        rows.append("</tbody></table>")
        if cfg["include_specs"]:
            rows.append(f"<div><strong>{label(cfg['lang'], 'specs_header')}</strong><p>{label(cfg['lang'], 'specs_placeholder')}</p></div>")
        sections.append("".join(rows))

    return "".join(sections)


def build_package_detail_preview(df: pd.DataFrame, cfg: dict) -> str:
    markup = cfg["markup"]
    tr = cfg["tax_rate"] / 100.0
    rows_data = all_items_with_subitems(df)
    suppliers = all_suppliers(df)
    show_toggle = cfg["include_toggle"]

    subtotals: dict[str, float] = {supplier: 0.0 for supplier in suppliers}
    for row in rows_data:
        for supplier in suppliers:
            base = price_for_any_type(df, supplier, row["code"], row["description"])
            if base > 0:
                subtotals[supplier] += marked_up_price(base, supplier, markup, row["description"])

    winner = min(
        (s for s in suppliers if subtotals.get(s, 0) > 0),
        key=lambda s: subtotals[s] * (1 + tr),
        default="",
    )

    header = ["", label(cfg["lang"], "brand"), label(cfg["lang"], "code"), label(cfg["lang"], "description"), label(cfg["lang"], "image")]
    if show_toggle:
        header.append(label(cfg["lang"], "incl_in_total"))
    header.append(label(cfg["lang"], "qty"))
    header.extend(suppliers)

    html = [preview_styles(), "<table class='pa-table'><thead><tr>"]
    for heading in header:
        cls = "winner-head" if heading == winner else ""
        html.append(f"<th class='{cls}'>{heading}</th>")
    html.append("</tr></thead><tbody>")

    for row in rows_data:
        is_sub = row["type"] == "subitem"
        html.append("<tr>")
        html.append(f"<td class='center'>{'↳' if is_sub else ''}</td>")
        html.append(f"<td>{'' if is_sub else row['brand']}</td>")
        html.append(f"<td>{'' if is_sub else row['code']}</td>")
        desc_cls = "subitem" if is_sub else ""
        html.append(f"<td class='{desc_cls}'>{row['description']}</td>")
        html.append("<td></td>")
        if show_toggle:
            html.append("<td class='center'>✓</td>")
        html.append("<td class='center'>1</td>")
        for supplier in suppliers:
            base = price_for_any_type(df, supplier, row["code"], row["description"])
            if base > 0:
                display = currency(marked_up_price(base, supplier, markup, row["description"]))
                cls = "num winner-col" if supplier == winner else "num"
            else:
                display = "—" if is_sub else ""
                cls = "center winner-col" if supplier == winner else "center"
            html.append(f"<td class='{cls}'>{display}</td>")
        html.append("</tr>")

    tax_totals = {s: subtotals[s] * tr for s in suppliers}
    grand_totals = {s: subtotals[s] + tax_totals[s] for s in suppliers}
    colspan = 7 if show_toggle else 6
    for key, source in [
        ("subtotal", subtotals),
        ("tax", tax_totals),
        ("total", grand_totals),
    ]:
        suffix = f" {cfg['tax_label']}" if key == "tax" else ""
        html.append(f"<tr><td colspan='{colspan}' class='summary-label'>{label(cfg['lang'], key)}{suffix}</td>")
        for supplier in suppliers:
            cls = "num winner-col" if supplier == winner else "num"
            html.append(f"<td class='{cls}'>{currency(source[supplier])}</td>")
        html.append("</tr>")

    html.append(f"<tr class='best-row'><td colspan='{colspan}'></td>")
    for supplier in suppliers:
        html.append(f"<td>{label(cfg['lang'], 'best_price') if supplier == winner else ''}</td>")
    html.append("</tr></tbody></table>")
    return "".join(html)


def reset_for_new_upload(file_name: str | None, file_bytes: bytes | None) -> None:
    current_sig = (file_name, file_bytes)
    if st.session_state.get("upload_signature") != current_sig:
        st.session_state.upload_signature = current_sig
        st.session_state.excel_data = None
        st.session_state.df = None


def validate_df(df: pd.DataFrame) -> list[str]:
    return [col for col in REQUIRED_COLUMNS if col not in df.columns]


for key, default in {
    "df": None,
    "excel_data": None,
    "upload_signature": None,
}.items():
    st.session_state.setdefault(key, default)

lang = st.sidebar.selectbox("en / fr", options=["en", "fr"], index=0)
mode = st.sidebar.radio(label(lang, "line_item"), options=["individual", "package_detail"], index=0)
province_options = list(PROVINCE_TAX.keys())
province = st.sidebar.selectbox(label(lang, "tax"), options=province_options, index=province_options.index("QC"))
include_toggle = st.sidebar.checkbox(label(lang, "incl_in_total"), value=True)
include_specs = st.sidebar.checkbox(
    label(lang, "specs_header"),
    value=False,
    disabled=mode == "package_detail",
)
if mode == "package_detail":
    include_specs = False

st.title(label(lang, "sheet_name"))
st.caption(f"{label(lang, 'default_title_ind') if mode == 'individual' else label(lang, 'default_title_pkg')} · {province}")

uploaded_file = st.file_uploader("CSV", type=["csv"])
file_bytes = uploaded_file.getvalue() if uploaded_file else None
reset_for_new_upload(uploaded_file.name if uploaded_file else None, file_bytes)

if uploaded_file and file_bytes is not None:
    df = normalize_df(pd.read_csv(uploaded_file))
    missing = validate_df(df)
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}")
    else:
        st.session_state.df = df
        st.success(f"{label(lang, 'sheet_name')}: {uploaded_file.name}")

if st.session_state.df is not None:
    st.subheader(label(lang, "description"))
    st.session_state.df = normalize_df(
        st.data_editor(st.session_state.df, use_container_width=True, num_rows="dynamic")
    )

    suppliers = all_suppliers(st.session_state.df)
    st.sidebar.markdown("---")
    st.sidebar.subheader(label(lang, "markup_pct"))
    markup = {
        supplier: st.sidebar.number_input(
            supplier,
            min_value=0.0,
            step=0.25,
            value=float(st.session_state.get(f"markup_{supplier}", 0.0)),
            key=f"markup_{supplier}",
        )
        for supplier in suppliers
    }

    cfg = {
        "lang": lang,
        "tax_rate": PROVINCE_TAX[province],
        "tax_label": f"{PROVINCE_TAX[province]}% {province}",
        "include_toggle": include_toggle,
        "include_specs": include_specs,
        "include_image": False,
        "markup": markup,
        "output_xlsx": "",
    }

    st.subheader("HTML preview")
    preview_html = build_individual_preview(st.session_state.df, cfg) if mode == "individual" else build_package_detail_preview(st.session_state.df, cfg)
    st.components.v1.html(preview_html, height=900, scrolling=True)

    if st.button("Generate Excel File"):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            cfg["output_xlsx"] = tmp.name
        if mode == "individual":
            build_individual(st.session_state.df, cfg)
        else:
            build_package_detail(st.session_state.df, cfg)
        with open(cfg["output_xlsx"], "rb") as fh:
            st.session_state.excel_data = fh.read()
        Path(cfg["output_xlsx"]).unlink(missing_ok=True)
        st.success("Excel file generated.")

if st.session_state.excel_data is not None:
    st.download_button(
        "Download Excel",
        data=st.session_state.excel_data,
        file_name="price_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
