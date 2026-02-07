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
    "Upload 1 or more PDF quotes",
    type=["pdf"],
    accept_multiple_files=True
)

if pdfs:  # at least one file uploaded
    if st.button("🚀 Process PDFs via Power Automate"):
        with st.spinner("Sending PDFs to Power Automate…"):

            files_payload = []
            for pdf in pdfs:
                encoded = base64.b64encode(pdf.read()).decode("ascii")
                files_payload.append({
                    "name": pdf.name,
                    "content": encoded
                })

            response = requests.post(
                POWER_AUTOMATE_URL,
                json={"files": files_payload},
                headers={
                    "Content-Type": "application/json"
                },
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
        
        # 🔥 Store CSV bytes for download
        st.session_state.csv_bytes = csv_bytes

        st.success(f"✅ CSV generated from {len(pdfs)} PDF(s) and loaded")
        st.rerun()
else:
    st.info("Upload 1 or more PDFs to start processing")

# 🔥 AUTO-DOWNLOAD CSV BUTTON (appears after processing)
if "csv_bytes" in st.session_state and st.session_state.csv_bytes is not None:
    st.download_button(
        label="📥 Download Raw CSV",
        data=st.session_state.csv_bytes,
        file_name=f"quotes_raw_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv"
    )

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
# GENERATE FINAL EXCEL (MINIMALIST FORMATTING)
# -------------------------------------------------
st.subheader("📥 Generate Final Excel")

if st.button("Generate Excel File"):
    df = st.session_state.df
    tax_rate = tax_percent / 100

    wb = Workbook()
    ws = wb.active
    ws.title = "Price Analysis"
    ws.sheet_view.showGridLines = False  # Clean minimalist look
    
    # --- MINIMALIST DESIGN TOKENS ---
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
    
    HEADER_BG = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
    WINNER_BG = PatternFill(start_color="F2FAF2", end_color="F2FAF2", fill_type="solid")
    SUBTLE_BORDER = Border(bottom=Side(style='thin', color="E5E5E5"))
    TEXT_PRIMARY = "1D1D1F"
    TEXT_SECONDARY = "86868B"

    current_row = 3  # Start with breathing room
    
    # === SINGLE OPTION TITLE AT TOP ===
    title_cell = ws.cell(row=1, column=2, value="Option 01")
    title_cell.font = Font(name='Arial', bold=True, size=14, color=TEXT_PRIMARY)

    main_items = df[
        (df["type"] == "item") & 
        df["Power Type"].notna() & 
        (df["Power Type"] != "")
    ]

    for opt_idx, (code, power_type) in enumerate(
        main_items[["code", "Power Type"]].drop_duplicates().values, 1
    ):
        # Get all items for this code/power type combination
        items_for_code = df[
            (df["code"] == code) &
            (
                (df["Power Type"] == power_type) |
                (df["Power Type"].isna()) |
                (df["Power Type"] == "")
            ) &
            (df["type"].isin(["item", "subitem"]))
        ]

        suppliers = list(items_for_code["supplier"].unique())
        brand = items_for_code[items_for_code["type"] == "item"].iloc[0]["brand"]
        descriptions = list(items_for_code["description"].unique())

        # --- DETERMINE WINNER (LOWEST TOTAL PRICE) ---
        winner_supplier = ""
        min_total = float('inf')
        for sup in suppliers:
            sup_items = items_for_code[items_for_code["supplier"] == sup]
            total = sup_items["price"].sum() * (1 + tax_rate)
            if total < min_total:
                min_total = total
                winner_supplier = sup


        # === 2. HEADER ROW ===
        header_row = current_row
        ws.row_dimensions[header_row].height = 28

        headers = ["DETAILS", "IMAGE", "QTY", "LINE ITEM"] + suppliers
        for i, h in enumerate(headers):
            col_idx = i + 2
            cell = ws.cell(row=header_row, column=col_idx, value=h.upper())
            cell.font = Font(name='Arial', size=8, bold=True, color=TEXT_SECONDARY)
            cell.fill = WINNER_BG if h == winner_supplier else HEADER_BG
            
            # Center supplier columns, left-align others
            if col_idx >= 6:  # Supplier columns
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(
                    horizontal="left" if col_idx <= 5 else "right",
                    vertical="center"
                )
            cell.border = SUBTLE_BORDER

        # === 3. CONTENT ROWS ===
        start_data_row = header_row + 1
        num_body_rows = len(descriptions) + 3  # +3 for total before tax, tax, and total rows

        # DETAILS column (merged)
        detail_val = f"BRAND\n{brand}\n\nCODE\n{code}\n\nPOWER\n{power_type}"
        d_cell = ws.cell(row=start_data_row, column=2, value=detail_val)
        d_cell.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        d_cell.font = Font(name='Arial', size=8, color=TEXT_SECONDARY)
        ws.merge_cells(
            start_row=start_data_row,
            start_column=2,
            end_row=start_data_row + num_body_rows,
            end_column=2
        )

        # IMAGE placeholder (merged)
        img_cell = ws.cell(row=start_data_row, column=3, value="[ PHOTO ]")
        img_cell.alignment = Alignment(horizontal="center", vertical="center")
        img_cell.font = Font(name='Arial', size=7, color="CCCCCC", italic=True)
        ws.merge_cells(
            start_row=start_data_row,
            start_column=3,
            end_row=start_data_row + num_body_rows,
            end_column=3
        )

        # ITEM ROWS
        for idx, desc in enumerate(descriptions):
            r_num = start_data_row + idx
            ws.row_dimensions[r_num].height = 32

            # QTY column
            qty_cell = ws.cell(row=r_num, column=4, value=1)
            qty_cell.alignment = Alignment(horizontal="center", vertical="center")
            qty_cell.font = Font(name='Arial', size=10, color=TEXT_PRIMARY)

            # LINE ITEM (description)
            desc_cell = ws.cell(row=r_num, column=5, value=desc)
            desc_cell.font = Font(name='Arial', size=10, color=TEXT_PRIMARY)
            desc_cell.alignment = Alignment(vertical="center")

            # SUPPLIER PRICES (with formulas tied to QTY)
            qty_letter = get_column_letter(4)
            for s_idx, sup in enumerate(suppliers):
                col = 6 + s_idx

                # Get price for this supplier/description combo
                price_row = items_for_code[
                    (items_for_code["supplier"] == sup) &
                    (items_for_code["description"] == desc)
                ]
                price = float(price_row["price"].iloc[0]) if not price_row.empty else 0

                # Formula: =QTY * price
                cell = ws.cell(row=r_num, column=col, value=f"={qty_letter}{r_num}*{price}")
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = SUBTLE_BORDER
                cell.font = Font(name='Arial', size=10, color=TEXT_PRIMARY)

                if sup == winner_supplier:
                    cell.fill = WINNER_BG

# === 4. TOTAL BEFORE TAX ROW ===
        total_before_tax_row = start_data_row + len(descriptions)
        ws.row_dimensions[total_before_tax_row].height = 28

        total_before_tax_label = ws.cell(row=total_before_tax_row, column=5, value="Total Before Tax")
        total_before_tax_label.font = Font(name='Arial', size=9, bold=True, color=TEXT_PRIMARY)
        total_before_tax_label.alignment = Alignment(vertical="center")

        for s_idx, sup in enumerate(suppliers):
            col = 6 + s_idx
            col_letter = get_column_letter(col)

            # Total Before Tax formula: SUM(items)
            tbt_cell = ws.cell(
                row=total_before_tax_row,
                column=col,
                value=f"=SUM({col_letter}{start_data_row}:{col_letter}{total_before_tax_row-1})"
            )
            tbt_cell.number_format = '#,##0.00'
            tbt_cell.alignment = Alignment(horizontal="right", vertical="center")
            tbt_cell.font = Font(name='Arial', size=9, bold=True, color=TEXT_PRIMARY)
            tbt_cell.border = SUBTLE_BORDER

            if sup == winner_supplier:
                tbt_cell.fill = WINNER_BG

        # === 5. TAX ROW ===
        tax_row = total_before_tax_row + 1
        ws.row_dimensions[tax_row].height = 28

        tax_label = ws.cell(row=tax_row, column=5, value=f"Tax ({int(tax_rate*100)}%)")
        tax_label.font = Font(name='Arial', size=9, color=TEXT_SECONDARY)
        tax_label.alignment = Alignment(vertical="center")

        for s_idx, sup in enumerate(suppliers):
            col = 6 + s_idx
            col_letter = get_column_letter(col)

            # Tax formula: Subtotal * tax_rate
            t_cell = ws.cell(
                row=tax_row,
                column=col,
                value=f"={col_letter}{total_before_tax_row}*{tax_rate}"
            )
            t_cell.number_format = '#,##0.00'
            t_cell.alignment = Alignment(horizontal="right", vertical="center")
            t_cell.font = Font(name='Arial', size=9, color=TEXT_SECONDARY)
            t_cell.border = SUBTLE_BORDER

            if sup == winner_supplier:
                t_cell.fill = WINNER_BG

        # === 6. TOTAL ROW ===
        total_row = tax_row + 1
        ws.row_dimensions[total_row].height = 40

        total_label = ws.cell(row=total_row, column=5, value="Total")
        total_label.font = Font(name='Arial', bold=True, size=11, color=TEXT_PRIMARY)
        total_label.alignment = Alignment(vertical="center")

        for s_idx, sup in enumerate(suppliers):
            col = 6 + s_idx
            col_letter = get_column_letter(col)

            # Total formula: Subtotal + Tax
            tot_cell = ws.cell(
                row=total_row,
                column=col,
                value=f"={col_letter}{total_before_tax_row}+{col_letter}{tax_row}"
            )
            tot_cell.font = Font(name='Arial', bold=True, size=11, color=TEXT_PRIMARY)
            tot_cell.number_format = '#,##0.00'
            tot_cell.alignment = Alignment(horizontal="right", vertical="center")
            tot_cell.border = Border(bottom=Side(style='medium', color="E5E5E5"))

            if sup == winner_supplier:
                tot_cell.fill = WINNER_BG


        # === 6. SPECS & DESCRIPTION BLOCK ===
        specs_row = total_row + 1
        ws.row_dimensions[specs_row].height = 60

        specs_label = ws.cell(row=specs_row, column=4, value="SPECS & DESCRIPTION")
        specs_label.font = Font(name='Arial', size=8, bold=True, color=TEXT_SECONDARY)
        specs_label.alignment = Alignment(vertical="top")

        specs_content = ws.cell(
            row=specs_row,
            column=5,
            value="Enter item specifications, dimensions, and technical details here..."
        )
        specs_content.font = Font(name='Arial', size=9, color=TEXT_SECONDARY, italic=True)
        specs_content.alignment = Alignment(wrap_text=True, vertical="top")

        # Merge across all supplier columns
        last_col = 5 + len(suppliers)
        ws.merge_cells(
            start_row=specs_row,
            start_column=5,
            end_row=specs_row,
            end_column=last_col
        )
        specs_content.border = Border(bottom=Side(style='thin', color="F2F2F2"))

        # Move to next table (with spacing)
        current_row = specs_row + 4

    # === COLUMN WIDTH ADJUSTMENTS ===
    ws.column_dimensions['B'].width = 16  # Details
    ws.column_dimensions['C'].width = 14  # Image
    ws.column_dimensions['D'].width = 8   # QTY
    ws.column_dimensions['E'].width = 30  # Line Item

    for i in range(len(suppliers)):
        ws.column_dimensions[get_column_letter(6+i)].width = 16

    # Save to downloadable file
    output = io.BytesIO()
    wb.save(output)

    st.download_button(
        "Download Excel",
        data=output.getvalue(),
        file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )






