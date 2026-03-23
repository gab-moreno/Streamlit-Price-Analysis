#!/usr/bin/env python3
"""
FF&E Price Analysis — Streamlit App
Delegates Excel generation to build_comparison.py builder functions.
"""

import streamlit as st
import pandas as pd
import tempfile
import os
import sys
from datetime import datetime

# ── Import builders and helpers from build_comparison ────────────────────────
_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)

from build_comparison import (
    build_individual,
    build_package_detail,
    PROVINCE_TAX,
    _LABELS,
    L,
    groups_from,
    items_for,
    uniq_col,
    first_brand,
    price_for,
    all_items_with_subitems,
    price_for_any_type,
    is_service,
)

st.set_page_config(layout="wide")

# ── Streamlit-only UI labels (not present in _LABELS) ────────────────────────
_UI = {
    "en": {
        "app_title":      "FF&E Price Analysis",
        "upload_csv":     "Upload CSV",
        "csv_loaded":     "CSV loaded",
        "review_table":   "Review Source Table",
        "preview_title":  "Price Analysis Preview",
        "generate_title": "Generate Excel",
        "generate_btn":   "Generate Excel File",
        "download_btn":   "Download Excel",
        "upload_prompt":  "Upload a CSV file to see the preview.",
        "mode_label":     "Mode",
        "mode_ind":       "Individual — per-item tables",
        "mode_pkg":       "Package Detail — single flat table",
        "province_label": "Province / Tax",
        "toggle_label":   "Include 'Incl. in Total' column",
        "specs_label":    "Include Specs block (Individual only)",
        "markup_header":  "Markup per Supplier (%)",
        "markup_sup":     "{sup}",
    },
    "fr": {
        "app_title":      "Analyse de prix FF&E",
        "upload_csv":     "Télécharger CSV",
        "csv_loaded":     "CSV chargé",
        "review_table":   "Réviser le tableau source",
        "preview_title":  "Aperçu de l'analyse de prix",
        "generate_title": "Générer le fichier Excel",
        "generate_btn":   "Générer le fichier Excel",
        "download_btn":   "Télécharger Excel",
        "upload_prompt":  "Téléchargez un fichier CSV pour voir l'aperçu.",
        "mode_label":     "Mode",
        "mode_ind":       "Individuel — tableaux par article",
        "mode_pkg":       "Détail forfait — tableau unique",
        "province_label": "Province / Taxe",
        "toggle_label":   "Inclure la colonne « Incl. au total »",
        "specs_label":    "Inclure les spécifications (individuel seulement)",
        "markup_header":  "Majoration par fournisseur (%)",
        "markup_sup":     "{sup}",
    },
}

# ── Session state ─────────────────────────────────────────────────────────────
for _k in ("df", "excel_data", "excel_filename"):
    if _k not in st.session_state:
        st.session_state[_k] = None

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    lang = st.selectbox(
        "Language / Langue",
        options=["en", "fr"],
        format_func=lambda x: "English" if x == "en" else "Français",
    )
    UI   = _UI[lang]
    _lc  = {"lang": lang}   # mini-cfg for L() calls

    mode = st.radio(
        UI["mode_label"],
        options=["individual", "package_detail"],
        format_func=lambda x: UI["mode_ind"] if x == "individual" else UI["mode_pkg"],
    )

    province_opts = list(PROVINCE_TAX.keys())
    province = st.selectbox(
        UI["province_label"],
        options=province_opts,
        index=province_opts.index("QC"),
        format_func=lambda p: f"{p} — {PROVINCE_TAX[p]}%",
    )
    tax_rate  = PROVINCE_TAX[province]
    tax_label = f"{tax_rate}% {province}"

    include_toggle = st.checkbox(UI["toggle_label"], value=True)
    include_specs  = st.checkbox(
        UI["specs_label"],
        value=False,
        disabled=(mode == "package_detail"),
    )

    # Per-supplier markup — shown once CSV is loaded
    markup = {}
    if st.session_state.df is not None:
        st.divider()
        st.subheader(UI["markup_header"])
        _df_sups = [
            s for s in st.session_state.df["supplier"].dropna().unique()
            if s and s != "nan"
        ]
        for sup in _df_sups:
            pct = st.number_input(
                UI["markup_sup"].format(sup=sup),
                min_value=0.0, max_value=100.0, value=0.0, step=1.0,
                key=f"mkp_{sup}",
            )
            markup[sup] = pct

# ── MAIN ──────────────────────────────────────────────────────────────────────
st.title(UI["app_title"])

# Upload
uploaded = st.file_uploader(UI["upload_csv"], type=["csv"])
if uploaded:
    _df = pd.read_csv(uploaded, keep_default_na=False)
    for _col in ["type", "supplier", "brand", "code", "description", "Power Type"]:
        if _col in _df.columns:
            _df[_col] = _df[_col].astype(str).str.strip()
    _df["price"] = pd.to_numeric(_df["price"], errors="coerce").fillna(0)
    st.session_state.df = _df.copy()
    st.session_state.excel_data = None   # reset stale Excel on new upload
    st.success(UI["csv_loaded"])

# Data editor
if st.session_state.df is not None:
    st.subheader(UI["review_table"])
    st.session_state.df = st.data_editor(
        st.session_state.df, use_container_width=True, num_rows="dynamic"
    )

# ── HTML PREVIEW ──────────────────────────────────────────────────────────────
_CSS = """<style>
.ffa{border-collapse:collapse;width:100%;margin-bottom:20px;
     font-family:'Segoe UI',Arial,sans-serif;font-size:12px;}
.ffa th,.ffa td{border:1px solid #E5E5E5;padding:5px 9px;vertical-align:middle;}
.ffa .hblue{background:#288AD6;color:#fff;font-size:13px;font-weight:700;
             text-align:left;padding-left:14px;border:none;}
.ffa th{background:#F8F9FA;color:#86868B;font-size:10px;font-weight:600;
         text-transform:uppercase;text-align:center;}
.ffa .det{background:#FAFAFA;color:#86868B;font-size:11px;line-height:1.7;}
.ffa .img{background:#fff;color:#ccc;text-align:center;font-style:italic;}
.ffa .win{background:#F2FAF2!important;}
.ffa .pr{text-align:right;}
.ffa .ctr{text-align:center;}
.ffa .tbf td{font-weight:bold;border-top:2px solid #E5E5E5!important;}
.ffa .txr td{color:#86868B;font-size:11px;}
.ffa .tot td{font-weight:bold;font-size:13px;border-top:2px solid #288AD6!important;}
.ffa .bst td{background:#F2FAF2;color:#1B5E20;font-weight:bold;text-align:center;}
.ffa .sub td{font-size:10px;color:#555;font-style:italic;background:#F0F4FA;}
.ffa .sub .win{background:#E8F5E9!important;}
.ffa-sep{height:4px;background:#E0E7EF;border-bottom:2px solid #288AD6;margin-bottom:0;}
</style>"""


def _hp(v):
    """Format a price value for HTML. Returns empty string for zero/None."""
    if not v:
        return ""
    return f"${v:,.2f}"


def _mk(base, pct, desc):
    """Apply markup percentage; services receive no markup."""
    if base <= 0:
        return 0.0
    return base if is_service(desc) else base * (1 + pct / 100.0)


def _winner(sups, totals, tax_rate):
    if not sups:
        return ""
    return min(sups, key=lambda s: totals[s] * (1 + tax_rate) if totals[s] > 0 else float("inf"))


# ── Individual mode HTML preview ──────────────────────────────────────────────
def html_individual(df, tax_rate, tax_label, markup, lang, include_toggle):
    tog   = include_toggle
    out   = [f"<div style='overflow-x:auto'>{_CSS}"]
    groups = groups_from(df)

    for gi, (code, pt) in enumerate(groups):
        itms  = items_for(df, code, pt)
        sups  = uniq_col(itms, "supplier")
        brand = first_brand(itms)
        descs = uniq_col(itms, "description")
        ns    = len(sups)
        ni    = len(descs)

        # Prices with markup applied
        dp = {d: {s: _mk(price_for(itms, s, d), markup.get(s, 0), d) for s in sups} for d in descs}

        totals = {s: sum(dp[d][s] for d in descs) for s in sups}
        taxes  = {s: totals[s] * tax_rate for s in sups}
        finals = {s: totals[s] + taxes[s] for s in sups}
        win    = _winner(sups, totals, tax_rate)

        # Column span constants
        # Visible cols: DETAILS + IMAGE + [INC] + QTY + LINE_ITEM + sups
        ncols = 4 + (1 if tog else 0) + ns
        # Within det/img rowspan: [INC] + QTY + LINE_ITEM
        lspan = 2 + (1 if tog else 0)
        # Outside rowspan (Total row): DETAILS + IMAGE + [INC] + QTY + LINE_ITEM
        ospan = 4 + (1 if tog else 0)
        # det/img rowspan covers: ni items + TBF + Tax
        rs = ni + 2

        if gi > 0:
            out.append("<div class='ffa-sep'></div>")

        out.append("<table class='ffa'>")

        # Group identifier header
        out.append(
            f"<tr><td colspan='{ncols}' class='hblue'>"
            f"{L('code_label', {'lang': lang})} {code} — {pt}</td></tr>"
        )

        # Column headers
        out.append("<tr>")
        out.append(f"<th style='text-align:left'>{L('details', {'lang': lang})}</th>")
        out.append(f"<th>{L('image', {'lang': lang})}</th>")
        if tog:
            out.append(f"<th>{L('incl_in_total', {'lang': lang})}</th>")
        out.append(f"<th>{L('qty', {'lang': lang})}</th>")
        out.append(f"<th style='text-align:left'>{L('line_item', {'lang': lang})}</th>")
        for s in sups:
            out.append(f"<th class='{'win' if s == win else ''}'>{s}</th>")
        out.append("</tr>")

        det_html = (
            f"<b>{L('brand_label', {'lang': lang})}</b><br>{brand}<br><br>"
            f"<b>{L('code_label', {'lang': lang})}</b><br>{code}<br><br>"
            f"<b>{L('power_label', {'lang': lang})}</b><br>{pt}"
        )

        # Item rows
        for i, desc in enumerate(descs):
            out.append("<tr>")
            if i == 0:
                out.append(f"<td rowspan='{rs}' class='det'>{det_html}</td>")
                out.append(f"<td rowspan='{rs}' class='img'>[ PHOTO ]</td>")
            if tog:
                out.append("<td class='ctr'>✓</td>")
            out.append("<td class='ctr' style='font-weight:bold'>1</td>")
            out.append(f"<td>{desc}</td>")
            for s in sups:
                v = dp[desc][s]
                out.append(
                    f"<td class='pr{' win' if s == win else ''}'>{_hp(v) if v > 0 else ''}</td>"
                )
            out.append("</tr>")

        # Total Before Tax (within det/img rowspan)
        out.append("<tr class='tbf'>")
        out.append(f"<td colspan='{lspan}'>{L('total_before_tax', {'lang': lang})}</td>")
        for s in sups:
            out.append(f"<td class='pr{' win' if s == win else ''}'>{_hp(totals[s])}</td>")
        out.append("</tr>")

        # Tax (within det/img rowspan — last row it covers)
        out.append("<tr class='txr'>")
        out.append(f"<td colspan='{lspan}'>{L('tax', {'lang': lang})} {tax_label}</td>")
        for s in sups:
            out.append(f"<td class='pr{' win' if s == win else ''}'>{_hp(taxes[s])}</td>")
        out.append("</tr>")

        # Total (outside rowspan — DETAILS & IMAGE cells are now free)
        out.append("<tr class='tot'>")
        out.append(f"<td colspan='{ospan}'></td>")
        for s in sups:
            out.append(f"<td class='pr{' win' if s == win else ''}'>{_hp(finals[s])}</td>")
        out.append("</tr>")

        out.append("</table>")

    out.append("</div>")
    return "\n".join(out)


# ── Package Detail mode HTML preview ─────────────────────────────────────────
def html_package_detail(df, tax_rate, tax_label, markup, lang, include_toggle):
    tog  = include_toggle
    rows = all_items_with_subitems(df)
    sups = [s for s in df["supplier"].dropna().unique() if s and s != "nan"]
    ns   = len(sups)

    # Prices with markup applied per row
    all_dp = [
        {s: _mk(price_for_any_type(df, s, r["code"], r["description"]),
                markup.get(s, 0), r["description"])
         for s in sups}
        for r in rows
    ]

    totals = {s: sum(dp[s] for dp in all_dp) for s in sups}
    taxes  = {s: totals[s] * tax_rate for s in sups}
    finals = {s: totals[s] + taxes[s] for s in sups}
    win    = _winner(sups, totals, tax_rate)

    # Columns: indent + BRAND + CODE + DESC + IMAGE + [INC] + QTY + sups
    ncols = 7 + (1 if tog else 0) + ns
    lspan = 7 + (1 if tog else 0)   # everything left of supplier columns

    out = [f"<div style='overflow-x:auto'>{_CSS}"]
    out.append("<table class='ffa'>")
    out.append(
        f"<tr><td colspan='{ncols}' class='hblue'>"
        f"{L('default_title_pkg', {'lang': lang})}</td></tr>"
    )

    # Column headers
    out.append("<tr>")
    out.append("<th></th>")   # indent
    out.append(f"<th style='text-align:left'>{L('brand', {'lang': lang})}</th>")
    out.append(f"<th style='text-align:left'>{L('code', {'lang': lang})}</th>")
    out.append(f"<th style='text-align:left'>{L('description', {'lang': lang})}</th>")
    out.append(f"<th>{L('image', {'lang': lang})}</th>")
    if tog:
        out.append(f"<th>{L('incl_in_total', {'lang': lang})}</th>")
    out.append(f"<th>{L('qty', {'lang': lang})}</th>")
    for s in sups:
        out.append(f"<th class='{'win' if s == win else ''}'>{s}</th>")
    out.append("</tr>")

    # Data rows
    item_counter = 0
    for ri, (row, dp) in enumerate(zip(rows, all_dp)):
        is_sub = row["type"] == "subitem"
        if not is_sub:
            item_counter += 1

        row_cls = "sub" if is_sub else ""
        out.append(f"<tr class='{row_cls}'>")

        # Indent marker
        out.append(
            f"<td class='ctr' style='color:#aaa;font-size:11px'>{'↳' if is_sub else ''}</td>"
        )

        # Brand & Code (blank for subitems)
        out.append(f"<td>{'  ' if is_sub else row['brand']}</td>")
        out.append(f"<td>{'  ' if is_sub else row['code']}</td>")

        # Description (indented for subitems)
        _ind = "padding-left:18px;" if is_sub else ""
        out.append(f"<td style='{_ind}'>{row['description']}</td>")

        # Image cell
        out.append("<td class='img' style='font-size:10px'>[ ]</td>")

        # Include toggle
        if tog:
            out.append("<td class='ctr'>✓</td>")

        # QTY
        out.append("<td class='ctr'>1</td>")

        # Supplier price cells
        for s in sups:
            v = dp[s]
            win_cls = " win" if s == win else ""
            if is_sub:
                display = _hp(v) if v > 0 else "—"
            else:
                display = _hp(v) if v > 0 else ""
            out.append(f"<td class='pr{win_cls}'>{display}</td>")

        out.append("</tr>")

    # Subtotal
    out.append("<tr class='tbf'>")
    out.append(
        f"<td colspan='{lspan}' style='text-align:right;padding-right:10px'>"
        f"{L('subtotal', {'lang': lang})}</td>"
    )
    for s in sups:
        out.append(f"<td class='pr{' win' if s == win else ''}'>{_hp(totals[s])}</td>")
    out.append("</tr>")

    # Tax
    out.append("<tr class='txr'>")
    out.append(
        f"<td colspan='{lspan}' style='text-align:right;padding-right:10px'>"
        f"{L('tax', {'lang': lang})} {tax_label}</td>"
    )
    for s in sups:
        out.append(f"<td class='pr{' win' if s == win else ''}'>{_hp(taxes[s])}</td>")
    out.append("</tr>")

    # Total
    out.append("<tr class='tot'>")
    out.append(
        f"<td colspan='{lspan}' style='text-align:right;padding-right:10px'>"
        f"{L('total', {'lang': lang})}</td>"
    )
    for s in sups:
        out.append(f"<td class='pr{' win' if s == win else ''}'>{_hp(finals[s])}</td>")
    out.append("</tr>")

    # Best Price banner
    out.append("<tr class='bst'>")
    out.append(f"<td colspan='{lspan}'></td>")
    for s in sups:
        out.append(f"<td>{L('best_price', {'lang': lang}) if s == win else ''}</td>")
    out.append("</tr>")

    out.append("</table></div>")
    return "\n".join(out)


# Render preview
st.subheader(UI["preview_title"])
if st.session_state.df is not None and not st.session_state.df.empty:
    if mode == "individual":
        _html = html_individual(
            st.session_state.df, tax_rate, tax_label, markup, lang, include_toggle
        )
    else:
        _html = html_package_detail(
            st.session_state.df, tax_rate, tax_label, markup, lang, include_toggle
        )
    st.markdown(_html, unsafe_allow_html=True)
else:
    st.info(UI["upload_prompt"])

# ── EXCEL GENERATION ──────────────────────────────────────────────────────────
st.subheader(UI["generate_title"])

if st.button(UI["generate_btn"]):
    if st.session_state.df is None:
        st.error("Please upload a CSV first.")
    else:
        cfg = {
            "lang":            lang,
            "tax_rate":        tax_rate,
            "tax_label":       tax_label,
            "include_toggle":  include_toggle,
            "include_specs":   include_specs if mode == "individual" else False,
            "include_image":   False,
            "markup":          markup,
        }

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as _tmp:
            _tmp_path = _tmp.name

        cfg["output_xlsx"] = _tmp_path

        try:
            if mode == "individual":
                build_individual(st.session_state.df, cfg)
            else:
                build_package_detail(st.session_state.df, cfg)

            with open(_tmp_path, "rb") as _f:
                st.session_state.excel_data = _f.read()

            st.session_state.excel_filename = (
                f"price_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
        except Exception as _e:
            st.error(f"Excel generation failed: {_e}")
            st.session_state.excel_data = None
        finally:
            if os.path.exists(_tmp_path):
                os.unlink(_tmp_path)

# Download button persists until next generation or new CSV upload
if st.session_state.excel_data is not None:
    st.download_button(
        UI["download_btn"],
        data=st.session_state.excel_data,
        file_name=st.session_state.excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
