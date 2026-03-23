import html
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
    items_for,
    price_for,
    price_for_any_type,
    uniq_col,
    is_service,
)

st.set_page_config(layout="wide")

REQUIRED_COLUMNS = ["type", "supplier", "brand", "code", "description", "Power Type", "price"]
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


EXTRA_LABELS = {
    "en": {
        "app_title": "Excel Summary",
        "language": "Language",
        "mode": "Mode",
        "mode_individual": "Individual",
        "mode_package_detail": "Package Detail",
        "province": "Province",
        "upload_csv": "Upload CSV",
        "loaded": "CSV loaded successfully.",
        "data_editor": "Review and edit data",
        "preview": "HTML preview",
        "generate": "Generate Excel File",
        "generated": "Excel file generated.",
        "download": "Download Excel File",
        "markup_section": "Markup % by supplier",
        "missing_columns": "Missing required columns",
        "no_preview": "Upload a CSV to preview the comparison table.",
        "empty_preview": "No preview rows are available for the selected mode.",
        "include_specs": "Include Specs block",
        "include_toggle": "Include Incl. in Total column",
    },
    "fr": {
        "app_title": "Résumé Excel",
        "language": "Langue",
        "mode": "Mode",
        "mode_individual": "Individuel",
        "mode_package_detail": "Détail du forfait",
        "province": "Province",
        "upload_csv": "Téléverser le CSV",
        "loaded": "CSV chargé avec succès.",
        "data_editor": "Vérifier et modifier les données",
        "preview": "Aperçu HTML",
        "generate": "Générer le fichier Excel",
        "generated": "Fichier Excel généré.",
        "download": "Télécharger le fichier Excel",
        "markup_section": "Majoration % par fournisseur",
        "missing_columns": "Colonnes requises manquantes",
        "no_preview": "Téléversez un CSV pour prévisualiser le tableau comparatif.",
        "empty_preview": "Aucune ligne d’aperçu n’est disponible pour le mode sélectionné.",
        "include_specs": "Inclure le bloc de spécifications",
        "include_toggle": "Inclure la colonne Incl. au total",
    },
}


def label(lang: str, key: str) -> str:
    return _LABELS.get(lang, _LABELS["en"]).get(key, _LABELS["en"].get(key, key))


def app_label(lang: str, key: str) -> str:
    return EXTRA_LABELS.get(lang, EXTRA_LABELS["en"]).get(key, EXTRA_LABELS["en"].get(key, key))


def text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if pd.isna(value):
        return ""
    value = str(value)
    return "" if value.lower() == "nan" else value.strip()


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    normalized = df.copy().fillna("")
    for col in ["type", "supplier", "brand", "code", "description", "Power Type"]:
        if col in normalized.columns:
            normalized[col] = normalized[col].map(text)
    if "price" in normalized.columns:
        normalized["price"] = pd.to_numeric(normalized["price"], errors="coerce").fillna(0.0)
    return normalized


def currency(value: float | int) -> str:
    return f"${float(value):,.2f}"


def render_text(value: object) -> str:
    return html.escape(text(value)).replace("\n", "<br>")


def marked_up_price(base_price: float, supplier: str, markup: dict[str, float], description: str) -> float:
    if base_price <= 0:
        return 0.0
    pct = markup.get(supplier, 0.0) / 100.0
    return base_price if is_service(text(description)) else base_price * (1 + pct)


def build_cfg(lang: str, mode: str, province: str, include_toggle: bool, include_specs: bool, markup: dict[str, float]) -> dict:
    tax_rate = PROVINCE_TAX[province]
    return {
        "lang": lang,
        "mode": mode,
        "tax_rate": tax_rate,
        "tax_label": f"{tax_rate}% {province}",
        "include_toggle": include_toggle,
        "include_specs": include_specs,
        "include_image": False,
        "markup": markup,
        "title": label(lang, "default_title_ind") if mode == "individual" else label(lang, "default_title_pkg"),
        "output_xlsx": "",
    }


def preview_styles() -> str:
    return """
    <style>
      .preview-wrap {font-family: Arial, sans-serif; color: #1d1d1f;}
      .preview-table {border-collapse: collapse; width: 100%; margin: 0 0 24px 0; table-layout: fixed;}
      .preview-table th, .preview-table td {border: 1px solid #d7dde5; padding: 8px 10px; vertical-align: middle; color: #1d1d1f;}
      .preview-table th {background: #f8f9fa; color: #666; font-size: 12px; letter-spacing: 0.03em; text-transform: uppercase;}
      .preview-title {margin: 0 0 10px 0; font-size: 20px; color: #1d1d1f;}
      .preview-subtitle {margin: 0 0 16px 0; font-size: 14px; color: #666;}
      .details-cell {background: #fafafa; white-space: pre-line; vertical-align: top; width: 180px;}
      .image-cell {background: #ffffff; width: 90px;}
      .qty-cell, .toggle-cell, .indent-cell {text-align: center;}
      .desc-cell {text-align: left;}
      .num {text-align: right; white-space: nowrap;}
      .winner-col, .winner-head {background: #f2faf2 !important;}
      .summary-label {font-weight: 700; text-align: right; background: #fbfbfc;}
      .section-gap {height: 18px;}
      .specs-block {padding: 12px 14px; background: #fafafa; border: 1px solid #e5e7eb; margin: 0 0 24px 0;}
      .subitem-row td {background: #f4f7fb; color: #555;}
      .subitem-desc {padding-left: 24px; font-style: italic;}
      .best-row td {background: #f2faf2; color: #1b5e20; font-weight: 700; text-align: center;}
      .muted {color: #9aa0a6;}
    </style>
    """


def supplier_totals_for_group(items: pd.DataFrame, suppliers: list[str], descriptions: list[str], markup: dict[str, float]) -> dict[str, float]:
    totals: dict[str, float] = {}
    for supplier in suppliers:
        total = 0.0
        for desc in descriptions:
            total += marked_up_price(price_for(items, supplier, desc), supplier, markup, desc)
        totals[supplier] = total
    return totals


def winner_from_totals(totals: dict[str, float], tax_rate: float) -> str:
    candidates = [supplier for supplier, total in totals.items() if total > 0]
    if not candidates:
        return ""
    return min(candidates, key=lambda supplier: totals[supplier] * (1 + tax_rate))


def build_individual_preview(df: pd.DataFrame, cfg: dict) -> str:
    groups = groups_from(df)
    if not groups:
        return ""

    lang = cfg["lang"]
    tax_rate = cfg["tax_rate"] / 100.0
    markup = cfg["markup"]
    html_parts = [preview_styles(), "<div class='preview-wrap'>"]

    for code, power_type in groups:
        items = items_for(df, code, power_type)
        suppliers = [text(supplier) for supplier in uniq_col(items, "supplier") if text(supplier)]
        descriptions = [text(desc) for desc in uniq_col(items, "description") if text(desc)]
        if not suppliers or not descriptions:
            continue

        brand = text(first_brand(items))
        totals_before_tax = supplier_totals_for_group(items, suppliers, descriptions, markup)
        winner = winner_from_totals(totals_before_tax, tax_rate)
        tax_totals = {supplier: total * tax_rate for supplier, total in totals_before_tax.items()}
        grand_totals = {supplier: totals_before_tax[supplier] + tax_totals[supplier] for supplier in suppliers}
        summary_colspan = 5 if cfg["include_toggle"] else 4
        merged_rowspan = len(descriptions)

        html_parts.append(f"<h3 class='preview-title'>{render_text(cfg['title'])}</h3>")
        html_parts.append(f"<div class='preview-subtitle'>{render_text(code)} — {render_text(power_type)}</div>")
        html_parts.append("<table class='preview-table'><thead><tr>")

        headers = [label(lang, "details"), label(lang, "image")]
        if cfg["include_toggle"]:
            headers.append(label(lang, "incl_in_total"))
        headers.extend([label(lang, "qty"), label(lang, "line_item")])
        headers.extend(suppliers)

        for header in headers:
            header_class = "winner-head" if header == winner else ""
            html_parts.append(f"<th class='{header_class}'>{render_text(header)}</th>")
        html_parts.append("</tr></thead><tbody>")

        details_text = (
            f"{label(lang, 'brand_label')}\n{brand}\n\n"
            f"{label(lang, 'code_label')}\n{code}\n\n"
            f"{label(lang, 'power_label')}\n{power_type}"
        )

        for index, desc in enumerate(descriptions):
            html_parts.append("<tr>")
            if index == 0:
                html_parts.append(f"<td class='details-cell' rowspan='{merged_rowspan}'>{render_text(details_text)}</td>")
                html_parts.append(f"<td class='image-cell' rowspan='{merged_rowspan}'></td>")
            if cfg["include_toggle"]:
                html_parts.append("<td class='toggle-cell'>✓</td>")
            html_parts.append("<td class='qty-cell'>1</td>")
            html_parts.append(f"<td class='desc-cell'>{render_text(desc)}</td>")
            for supplier in suppliers:
                base = price_for(items, supplier, desc)
                display = currency(marked_up_price(base, supplier, markup, desc)) if base > 0 else ""
                col_class = "num winner-col" if supplier == winner else "num"
                html_parts.append(f"<td class='{col_class}'>{display}</td>")
            html_parts.append("</tr>")

        summaries = [
            (label(lang, "total_before_tax"), totals_before_tax),
            (f"{label(lang, 'tax')} {cfg['tax_label']}", tax_totals),
            (label(lang, "total"), grand_totals),
        ]
        for summary_label, values in summaries:
            html_parts.append("<tr>")
            html_parts.append(f"<td colspan='{summary_colspan}' class='summary-label'>{render_text(summary_label)}</td>")
            for supplier in suppliers:
                col_class = "num winner-col" if supplier == winner else "num"
                html_parts.append(f"<td class='{col_class}'>{currency(values[supplier])}</td>")
            html_parts.append("</tr>")

        html_parts.append("</tbody></table>")
        if cfg["include_specs"]:
            html_parts.append(
                f"<div class='specs-block'><strong>{render_text(label(lang, 'specs_header'))}</strong><br><br>{render_text(label(lang, 'specs_placeholder'))}</div>"
            )
        html_parts.append("<div class='section-gap'></div>")

    html_parts.append("</div>")
    return "".join(html_parts)


def build_package_detail_preview(df: pd.DataFrame, cfg: dict) -> str:
    rows_data = all_items_with_subitems(df)
    suppliers = [text(supplier) for supplier in all_suppliers(df) if text(supplier)]
    if not rows_data or not suppliers:
        return ""

    lang = cfg["lang"]
    tax_rate = cfg["tax_rate"] / 100.0
    markup = cfg["markup"]
    subtotals = {supplier: 0.0 for supplier in suppliers}

    for row in rows_data:
        for supplier in suppliers:
            base = price_for_any_type(df, supplier, row["code"], row["description"])
            subtotals[supplier] += marked_up_price(base, supplier, markup, row["description"])

    winner = winner_from_totals(subtotals, tax_rate)
    tax_totals = {supplier: total * tax_rate for supplier, total in subtotals.items()}
    grand_totals = {supplier: subtotals[supplier] + tax_totals[supplier] for supplier in suppliers}
    summary_colspan = 7 if cfg["include_toggle"] else 6

    html_parts = [preview_styles(), "<div class='preview-wrap'>"]
    html_parts.append(f"<h3 class='preview-title'>{render_text(cfg['title'])}</h3>")
    html_parts.append("<table class='preview-table'><thead><tr>")

    headers = ["", label(lang, "brand"), label(lang, "code"), label(lang, "description"), label(lang, "image")]
    if cfg["include_toggle"]:
        headers.append(label(lang, "incl_in_total"))
    headers.append(label(lang, "qty"))
    headers.extend(suppliers)

    for header in headers:
        header_class = "winner-head" if header == winner else ""
        html_parts.append(f"<th class='{header_class}'>{render_text(header)}</th>")
    html_parts.append("</tr></thead><tbody>")

    for row in rows_data:
        is_subitem = row["type"] == "subitem"
        row_class = "subitem-row" if is_subitem else ""
        html_parts.append(f"<tr class='{row_class}'>")
        html_parts.append(f"<td class='indent-cell'>{'↳' if is_subitem else ''}</td>")
        html_parts.append(f"<td>{'' if is_subitem else render_text(row['brand'])}</td>")
        html_parts.append(f"<td>{'' if is_subitem else render_text(row['code'])}</td>")
        desc_class = "desc-cell subitem-desc" if is_subitem else "desc-cell"
        html_parts.append(f"<td class='{desc_class}'>{render_text(row['description'])}</td>")
        html_parts.append("<td></td>")
        if cfg["include_toggle"]:
            html_parts.append("<td class='toggle-cell'>✓</td>")
        html_parts.append("<td class='qty-cell'>1</td>")
        for supplier in suppliers:
            base = price_for_any_type(df, supplier, row["code"], row["description"])
            if base > 0:
                value = currency(marked_up_price(base, supplier, markup, row["description"]))
                col_class = "num winner-col" if supplier == winner else "num"
            else:
                value = "—" if is_subitem else ""
                col_class = "winner-col muted" if supplier == winner else "muted"
            html_parts.append(f"<td class='{col_class}'>{value}</td>")
        html_parts.append("</tr>")

    summaries = [
        (label(lang, "subtotal"), subtotals),
        (f"{label(lang, 'tax')} {cfg['tax_label']}", tax_totals),
        (label(lang, "total"), grand_totals),
    ]
    for summary_label, values in summaries:
        html_parts.append("<tr>")
        html_parts.append(f"<td colspan='{summary_colspan}' class='summary-label'>{render_text(summary_label)}</td>")
        for supplier in suppliers:
            col_class = "num winner-col" if supplier == winner else "num"
            html_parts.append(f"<td class='{col_class}'>{currency(values[supplier])}</td>")
        html_parts.append("</tr>")

    html_parts.append(f"<tr class='best-row'><td colspan='{summary_colspan}'></td>")
    for supplier in suppliers:
        html_parts.append(f"<td>{render_text(label(lang, 'best_price')) if supplier == winner else ''}</td>")
    html_parts.append("</tr></tbody></table></div>")
    return "".join(html_parts)


def validate_df(df: pd.DataFrame) -> list[str]:
    return [column for column in REQUIRED_COLUMNS if column not in df.columns]


def reset_for_new_upload(file_name: str | None, file_bytes: bytes | None) -> None:
    signature = (file_name, file_bytes)
    if st.session_state.get("upload_signature") != signature:
        st.session_state.upload_signature = signature
        st.session_state.df = None
        st.session_state.excel_data = None


for key, default in {
    "df": None,
    "excel_data": None,
    "upload_signature": None,
}.items():
    st.session_state.setdefault(key, default)

lang = st.sidebar.selectbox(app_label("en", "language"), options=["en", "fr"], format_func=lambda value: value.upper())
mode = st.sidebar.radio(
    app_label(lang, "mode"),
    options=["individual", "package_detail"],
    format_func=lambda value: app_label(lang, f"mode_{value}"),
)
province_codes = list(PROVINCE_TAX.keys())
province = st.sidebar.selectbox(app_label(lang, "province"), options=province_codes, index=province_codes.index("QC"))
include_toggle = st.sidebar.checkbox(app_label(lang, "include_toggle"), value=True)
include_specs = st.sidebar.checkbox(
    app_label(lang, "include_specs"),
    value=False,
    disabled=mode == "package_detail",
)
if mode == "package_detail":
    include_specs = False

st.title(app_label(lang, "app_title"))
st.caption(f"{label(lang, 'sheet_name')} · {province}")

uploaded_file = st.file_uploader(app_label(lang, "upload_csv"), type=["csv"])
file_bytes = uploaded_file.getvalue() if uploaded_file else None
reset_for_new_upload(uploaded_file.name if uploaded_file else None, file_bytes)

if uploaded_file and file_bytes is not None:
    candidate_df = normalize_df(pd.read_csv(uploaded_file))
    missing_columns = validate_df(candidate_df)
    if missing_columns:
        st.error(f"{app_label(lang, 'missing_columns')}: {', '.join(missing_columns)}")
    else:
        st.session_state.df = candidate_df
        st.success(app_label(lang, "loaded"))

if st.session_state.df is not None:
    st.subheader(app_label(lang, "data_editor"))
    st.session_state.df = normalize_df(
        st.data_editor(st.session_state.df, use_container_width=True, num_rows="dynamic")
    )

    suppliers = [text(supplier) for supplier in all_suppliers(st.session_state.df) if text(supplier)]
    st.sidebar.markdown("---")
    st.sidebar.subheader(app_label(lang, "markup_section"))
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

    cfg = build_cfg(lang, mode, province, include_toggle, include_specs, markup)
    st.subheader(app_label(lang, "preview"))
    preview_html = build_individual_preview(st.session_state.df, cfg) if mode == "individual" else build_package_detail_preview(st.session_state.df, cfg)
    if preview_html:
        st.markdown(preview_html, unsafe_allow_html=True)
    else:
        st.info(app_label(lang, "empty_preview"))

    if st.button(app_label(lang, "generate")):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            output_path = tmp.name
        cfg["output_xlsx"] = output_path
        try:
            if mode == "individual":
                build_individual(st.session_state.df, cfg)
            else:
                build_package_detail(st.session_state.df, cfg)
            st.session_state.excel_data = Path(output_path).read_bytes()
            st.success(app_label(lang, "generated"))
        finally:
            Path(output_path).unlink(missing_ok=True)
else:
    st.info(app_label(lang, "no_preview"))

if st.session_state.excel_data is not None:
    st.download_button(
        app_label(lang, "download"),
        data=st.session_state.excel_data,
        file_name="price_analysis.xlsx",
        mime=EXCEL_MIME,
    )
