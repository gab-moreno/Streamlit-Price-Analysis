import requests
import base64
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, numbers, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io



st.set_page_config(layout="wide")

POWER_AUTOMATE_URL = st.secrets["power_automate"]["url"]

# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------
if "df" not in st.session_state:
    st.session_state.df = None

# -------------------------------------------------
# HEADER
# -------------------------------------------------
st.title("📊 Interactive Table Review & Price Analysis")


# -------------------------------------------------
# PDF UPLOAD → POWER AUTOMATE (HIGHEST PRIORITY)
# -------------------------------------------------
st.subheader("📄 Upload 3 Quote PDFs (Power Automate)")

pdfs = st.file_uploader(
    "Upload exactly 3 PDF quotes",
    type=["pdf"],
    accept_multiple_files=True
)

if pdfs and len(pdfs) != 3:
    st.warning("Please upload exactly 3 PDF files.")

if pdfs and len(pdfs) == 3:
    if st.button("🚀 Process PDFs via Power Automate"):
        with st.spinner("Sending PDFs to Power Automate…"):

            files_payload = []
            for pdf in pdfs:
                encoded = base64.b64encode(pdf.read()).decode("utf-8")
                files_payload.append({
                    "name": pdf.name,
                    "content": encoded
                })

            response = requests.post(
                POWER_AUTOMATE_URL,
                json={"files": files_payload},
                timeout=180
            )

        if response.status_code != 200:
            st.error("Power Automate failed to process PDFs")
            st.stop()

        # Expecting base64 CSV back
        csv_bytes = base64.b64decode(response.json()["csv"])
        df = pd.read_csv(io.BytesIO(csv_bytes))

        # 🔑 HANDOFF POINT — everything else already works
        for col in ["type", "supplier", "brand", "code", "description", "Power Type"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()

        st.session_state.df = df
        st.session_state.current_job_path = None
        st.session_state.job_loaded_from_queue = False

        st.success("✅ CSV generated from PDFs and loaded")
        st.rerun()


# -------------------------------------------------
# UPLOAD FILE (MANUAL OVERRIDE)
# -------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload CSV or Excel (manual override)",
    type=["csv", "xlsx"]
)

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    for col in ["type", "supplier", "brand", "code", "description", "Power Type"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # 🔹 Override queue state
    st.session_state.df = df.copy()
    st.session_state.current_job_path = None
    st.session_state.job_loaded_from_queue = False

    st.success("📤 Manual file loaded (queue overridden)")



# -------------------------------------------------
# EDIT SOURCE TABLE
# -------------------------------------------------
if st.session_state.df is not None:
    st.subheader("✏️ Review Source Table")
    st.session_state.df = st.data_editor(
        st.session_state.df,
        use_container_width=True,
        num_rows="dynamic"
    )

# -------------------------------------------------
# TAX INPUT
# -------------------------------------------------
st.subheader("💲 Tax Settings")
tax_percent = st.number_input("Tax Percentage", min_value=0.0, value=12.0)

# -------------------------------------------------
# HTML PREVIEW (EXCEL-STYLE)
# -------------------------------------------------
st.subheader("👀 Price Analysis Preview (HTML Table)")

def generate_html_table(df, tax_percent):
    tax_rate = tax_percent / 100

    html = """
    <div style="overflow-x:auto;">
    <style>
        table {
            border-collapse: collapse !important;
            width: 100%;
            margin-bottom: 40px;
            font-family: Arial, sans-serif;
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #bfbfbf;
        }

        th, td {
            border: 1px solid #bfbfbf !important;
            padding: 6px 8px;
            vertical-align: middle;
            text-align: left;
            background-clip: padding-box;
        }

        th {
            background-color: #dae9f8;
            font-weight: 600;
        }

        .total-row td {
            background-color: #fce4d6;
            font-weight: bold;
        }
    </style>
    """

    main_items = df[
        (df["type"] == "item") &
        df["Power Type"].notna() &
        (df["Power Type"] != "")
    ]

    for code, power_type in main_items[["code", "Power Type"]].drop_duplicates().values:

        items_for_code = df[
            (df["code"] == code) &
            (
                (df["Power Type"] == power_type) |
                (df["Power Type"].isna()) |
                (df["Power Type"] == "")
            ) &
            (df["type"].isin(["item", "subitem"]))
        ]

        suppliers = items_for_code["supplier"].unique()
        brand = items_for_code[items_for_code["type"] == "item"].iloc[0]["brand"]
        descriptions = items_for_code["description"].unique()

        body_rows = len(descriptions) + 2  # items + tax + total

        html += "<table>"

        # HEADER
        html += "<tr>"
        html += "<th>Details</th><th></th><th>QTY</th><th>Items</th>"
        for s in suppliers:
            html += f"<th>{s}</th>"
        html += "</tr>"

        totals = {s: 0 for s in suppliers}

        # FIRST ITEM ROW (with DETAILS)
        first_desc = descriptions[0]

        html += "<tr>"
        html += f"""
            <td rowspan="{body_rows}">
                <b>Brand</b><br>{brand}<br><br>
                <b>Code</b><br>{code}<br><br>
                <b>Power Type</b><br>{power_type}
            </td>
            <td rowspan="{body_rows}"></td>
            <td>1</td>
            <td>{first_desc}</td>
        """

        for s in suppliers:
            row = items_for_code[
                (items_for_code["supplier"] == s) &
                (items_for_code["description"] == first_desc)
            ]
            price = float(row["price"].iloc[0]) if not row.empty else 0
            totals[s] += price
            html += f"<td>${price:,.2f}</td>"

        html += "</tr>"

        # REMAINING ITEM ROWS
        for desc in descriptions[1:]:
            html += "<tr>"
            html += f"<td>1</td><td>{desc}</td>"

            for s in suppliers:
                row = items_for_code[
                    (items_for_code["supplier"] == s) &
                    (items_for_code["description"] == desc)
                ]
                price = float(row["price"].iloc[0]) if not row.empty else 0
                totals[s] += price
                html += f"<td>${price:,.2f}</td>"

            html += "</tr>"

        # TAX ROW
        html += "<tr>"
        html += "<td></td><td><b>Tax</b></td>"
        for _ in suppliers:
            html += f"<td>{tax_percent:.2f}%</td>"
        html += "</tr>"

        # TOTAL ROW
        html += "<tr class='total-row'>"
        html += "<td></td><td>Total</td>"
        for s in suppliers:
            total = totals[s] * (1 + tax_rate)
            html += f"<td>${total:,.2f}</td>"
        html += "</tr>"

        html += "</table>"

    html += "</div>"
    return html


# 🔥 RENDER HTML (LIVE, REACTIVE)
if (
    "df" in st.session_state
    and st.session_state.df is not None
    and not st.session_state.df.empty
):
    html = generate_html_table(st.session_state.df, tax_percent)
    st.markdown(html, unsafe_allow_html=True)
else:
    st.info("⬆️ Upload or generate data to see the price analysis preview.")

# -------------------------------------------------
# GENERATE FINAL EXCEL (PROVEN FORMATTING)
# -------------------------------------------------
st.subheader("📥 Generate Final Excel")

if st.button("Generate Excel File"):
    df = st.session_state.df
    tax_rate = tax_percent / 100

    wb = Workbook()
    ws = wb.active
    ws.title = "Items Summary"

    header_fill = PatternFill(start_color="DAE9F8", end_color="DAE9F8", fill_type="solid")
    total_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    thin_side = Side(border_style="thin", color="000000")

    start_row_offset = 1
    start_col_offset = 1
    current_row = 1

    main_items = df[(df["type"] == "item") & df["Power Type"].notna() & (df["Power Type"] != "")]

    for code, power_type in main_items[["code", "Power Type"]].drop_duplicates().values:

        items_for_code = df[
            (df["code"] == code) &
            (
                (df["Power Type"] == power_type) |
                (df["Power Type"].isna()) |
                (df["Power Type"] == "")
            ) &
            (df["type"].isin(["item", "subitem"]))
        ]

        suppliers = items_for_code["supplier"].unique()
        brand = items_for_code[items_for_code["type"] == "item"].iloc[0]["brand"]
        descriptions = items_for_code["description"].unique()

        start_row = current_row
        data_row = start_row + 1

        ws.cell(row=start_row + start_row_offset, column=1 + start_col_offset, value="Details")
        ws.cell(row=start_row + start_row_offset, column=3 + start_col_offset, value="Image")
        ws.cell(row=start_row + start_row_offset, column=4 + start_col_offset, value="QTY")
        ws.cell(row=start_row + start_row_offset, column=5 + start_col_offset, value="Items")

        for i, supplier in enumerate(suppliers):
            ws.cell(
                row=start_row + start_row_offset,
                column=6 + i + start_col_offset,
                value=supplier
            )

        last_header_col = 5 + len(suppliers) + start_col_offset
        for col in range(1 + start_col_offset, last_header_col + 1):
            ws.cell(row=start_row + start_row_offset, column=col).fill = header_fill

        ws.cell(row=data_row + start_row_offset, column=1 + start_col_offset, value="Brand")
        ws.cell(row=data_row + start_row_offset, column=2 + start_col_offset, value=brand)

        ws.cell(row=data_row + 1 + start_row_offset, column=1 + start_col_offset, value="Code")
        ws.cell(row=data_row + 1 + start_row_offset, column=2 + start_col_offset, value=code)

        ws.cell(row=data_row + 2 + start_row_offset, column=1 + start_col_offset, value="Power Type")
        ws.cell(row=data_row + 2 + start_row_offset, column=2 + start_col_offset, value=power_type)

        for i_desc, desc in enumerate(descriptions):
            row = data_row + i_desc + start_row_offset
            ws.cell(row=row, column=4 + start_col_offset, value=1)
            ws.cell(row=row, column=5 + start_col_offset, value=desc)

            qty_letter = get_column_letter(4 + start_col_offset)

            for i, supplier in enumerate(suppliers):
                col_idx = 6 + i + start_col_offset
                price_row = items_for_code[
                    (items_for_code["supplier"] == supplier) &
                    (items_for_code["description"] == desc)
                ]
                price = float(price_row["price"].iloc[0]) if not price_row.empty else 0
                ws.cell(row=row, column=col_idx, value=f"={qty_letter}{row}*{price}")
                ws.cell(row=row, column=col_idx).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

        extra_rows = 2 if "subitem" not in items_for_code["type"].values else 0

        tax_row = data_row + len(descriptions) + extra_rows + start_row_offset
        ws.cell(row=tax_row, column=5 + start_col_offset, value="Tax")

        for i in range(len(suppliers)):
            col_idx = 6 + i + start_col_offset
            ws.cell(row=tax_row, column=col_idx, value=tax_rate)
            ws.cell(row=tax_row, column=col_idx).number_format = numbers.FORMAT_PERCENTAGE_00

        total_row = tax_row + 1
        ws.cell(row=total_row, column=5 + start_col_offset, value="Total").fill = total_fill

        first_item_row = data_row + start_row_offset
        last_item_row = tax_row - 1

        for i in range(len(suppliers)):
            col_idx = 6 + i + start_col_offset
            col_letter = get_column_letter(col_idx)
            ws.cell(
                row=total_row,
                column=col_idx,
                value=f"=SUM({col_letter}{first_item_row}:{col_letter}{last_item_row})*(1+{col_letter}{tax_row})"
            )
            ws.cell(row=total_row, column=col_idx).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            ws.cell(row=total_row, column=col_idx).fill = total_fill

        first_row = start_row + start_row_offset
        last_row = total_row
        first_col = 1 + start_col_offset
        last_col = 5 + len(suppliers) + start_col_offset

        for r in range(first_row, last_row + 1):
            for c in range(first_col, last_col + 1):
                ws.cell(row=r, column=c).border = Border(
                    top=thin_side if r == first_row else None,
                    bottom=thin_side if r == last_row else None,
                    left=thin_side if c == first_col else None,
                    right=thin_side if c == last_col else None,
                )

        current_row = total_row + 3

    output = io.BytesIO()
    wb.save(output)

    st.download_button(
        "Download Excel",
        data=output.getvalue(),
        file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

