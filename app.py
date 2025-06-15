import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io

st.set_page_config(page_title="Cleaning Validation App", layout="wide")
st.title("Cleaning Validation Protocol – MACO, Swab & Rinse Limits")

# ----------- 1. Data Input -----------
st.markdown("""
Upload the following files **or use the example files**:
- product_details.xlsx
- analytical_method_validation.xlsx
- solubility_cleaning.xlsx
- equipment_details.xlsx
- rating_criteria.xlsx
""")

mode = st.radio(
    "How do you want to provide data?",
    ("Use example files from repo", "Upload my own files")
)

def read_excel_or_none(f, **kwargs):
    try:
        return pd.read_excel(f, **kwargs)
    except Exception as e:
        st.error(f"Error loading {f}: {e}")
        return None

if mode == "Upload my own files":
    uploaded_details = st.file_uploader("Product Details", type=["xlsx"])
    uploaded_amv = st.file_uploader("Analytical Method Validation", type=["xlsx"])
    uploaded_solclean = st.file_uploader("Solubility & Cleaning", type=["xlsx"])
    uploaded_equips = st.file_uploader("Equipment Details", type=["xlsx"])
    uploaded_criteria = st.file_uploader("Rating Criteria (4 sheets)", type=["xlsx"])
else:
    uploaded_details = "product_details.xlsx"
    uploaded_amv = "analytical_method_validation.xlsx"
    uploaded_solclean = "solubility_cleaning.xlsx"
    uploaded_equips = "equipment_details.xlsx"
    uploaded_criteria = "rating_criteria.xlsx"

files_ready = all([uploaded_details, uploaded_amv, uploaded_solclean, uploaded_equips, uploaded_criteria])

if files_ready:
    df = read_excel_or_none(uploaded_details)
    df_equip = read_excel_or_none(uploaded_equips)
    df_amv = read_excel_or_none(uploaded_amv)
    df_solclean = read_excel_or_none(uploaded_solclean)
    templates = pd.read_excel(uploaded_criteria, sheet_name=None)

    if None in [df, df_equip, df_amv, df_solclean, templates]:
        st.stop()

    solubility_template = templates['Solubility']
    dose_template = templates['Dose']
    toxicity_template = templates['Toxicity']
    cleaning_template = templates['Cleaning']

    # ----------- 2. Group Assignments -----------
    def assign_solubility_group(sol, template):
        if pd.isna(sol): return None
        sol = str(sol).strip().lower()
        for _, row in template.iterrows():
            desc = str(row['Description']).strip().lower()
            if sol == desc:
                return row['Group']
        return None
    def assign_range_group(value, template):
        try: value = float(value)
        except: return None
        for _, row in template.iterrows():
            if row['Min'] <= value <= row['Max']:
                return row['Group']
        return None
    def assign_cleaning_group(val, template):
        if pd.isna(val): return None
        val = str(val).strip().lower()
        for _, row in template.iterrows():
            desc = str(row['Description']).strip().lower()
            if val == desc:
                return row['Group']
        return None

    df['Solubility_Group'] = df['Solubility'].apply(lambda x: assign_solubility_group(x, solubility_template))
    df['Dose_Group'] = df['Min Dose (mg)'].apply(lambda x: assign_range_group(x, dose_template))
    df['Toxicity_Group'] = df['ADE/PDE (µg/day)'].apply(lambda x: assign_range_group(x, toxicity_template))
    df['Cleaning_Group'] = df['Hardest To Clean'].apply(lambda x: assign_cleaning_group(x, cleaning_template))
    df['Worst_Case_Rating'] = (
        df['Solubility_Group'].astype(float) *
        df['Dose_Group'].astype(float) *
        df['Toxicity_Group'].astype(float) *
        df['Cleaning_Group'].astype(float)
    )
    df['BatchSize_Dose_Ratio'] = df['Min Batch Size (kg)'] / df['Max Dose (mg)']

    # ----------- 3. Worst Case Identification -----------
    prev_worst_case = df.loc[df['Worst_Case_Rating'].idxmax()]
    next_worst_case = df.loc[df['BatchSize_Dose_Ratio'].idxmin()]
    min_batch_next_kg = next_worst_case['Min Batch Size (kg)']
    max_dose_next_mg = next_worst_case['Max Dose (mg)']
    min_dose_prev_mg = prev_worst_case['Min Dose (mg)']
    ade_prev_ug = prev_worst_case['ADE/PDE (µg/day)']
    ade_prev_mg = ade_prev_ug / 1000

    # ----------- 4. MACO Calculations -----------
    maco_10ppm = 0.00001 * min_batch_next_kg * 1e6 / max_dose_next_mg
    maco_tdd = min_dose_prev_mg * min_batch_next_kg * 1e6 / (max_dose_next_mg * 1000)
    maco_ade = ade_prev_mg * min_batch_next_kg * 1e6 / max_dose_next_mg
    lowest_maco = min(maco_10ppm, maco_tdd, maco_ade)

    # ----------- 5. Surface Area -----------
    total_surface_area = df_equip['Product contact Surface Area (m2)'].sum()
    total_surface_area_with_margin = total_surface_area * 1.2

    # ----------- 6. Swab & Rinse Limits -----------
    swab_surface = df['Swab Surface in M. Sq.'].iloc[0]
    swab_limit = lowest_maco * swab_surface / total_surface_area_with_margin

    rinse_limits = []
    for idx, row in df_equip.iterrows():
        eq_surface = row['Product contact Surface Area (m2)']
        rinse_limit = lowest_maco * eq_surface / total_surface_area_with_margin
        rinse_vol = rinse_limit / 10
        rinse_limits.append({
            'Eq. Name': row['Eq. Name'],
            'Eq. ID': row['Eq. ID'],
            'Surface Area (m2)': eq_surface,
            'Rinse Limit (mg)': round(rinse_limit, 6),
            'Rinse Volume (L)': round(rinse_vol, 6),
            'Rinse Volume (ml)': round(rinse_vol * 1000, 2)
        })
    df_rinse_limits = pd.DataFrame(rinse_limits)

    # ----------- 7. Results Tabs -----------
    tab1, tab2, tab3, tab4 = st.tabs([
        "Product Group Assignment",
        "Worst Case Products",
        "MACO/Swab/Rinse Results",
        "Equipment & Rinse Table"
    ])
    with tab1:
        st.dataframe(df.head(20))
    with tab2:
        st.write("**Previous Worst Case (By Highest Rating):**")
        st.dataframe(prev_worst_case.to_frame().T)
        st.write("**Next Worst Case (By Lowest Min Batch / Max Dose ratio):**")
        st.dataframe(next_worst_case.to_frame().T)
    with tab3:
        st.write("**MACO Calculation (mg):**")
        st.write(f"10 ppm: `{maco_10ppm:.4f}` mg, TDD: `{maco_tdd:.4f}` mg, ADE/PDE: `{maco_ade:.4f}` mg, **Lowest Used**: `{lowest_maco:.4f}` mg")
        st.write(f"**Swab Limit:** `{swab_limit:.6f}` mg")
        st.write(f"**Total Equip Surface (with 20% margin):** {total_surface_area_with_margin:.2f} m²")
    with tab4:
        st.dataframe(df_rinse_limits)

    # ----------- 8. Download Excel Output -----------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Products with Groups")
        df_equip.to_excel(writer, index=False, sheet_name="Equipment")
        df_rinse_limits.to_excel(writer, index=False, sheet_name="Rinse Limits")
        pd.DataFrame([{
            "MACO_10ppm (mg)": maco_10ppm,
            "MACO_TDD (mg)": maco_tdd,
            "MACO_ADE (mg)": maco_ade,
            "Lowest MACO Used (mg)": lowest_maco,
            "Swab Limit (mg)": swab_limit,
            "Total Equip Surface (m2)": total_surface_area,
            "Total Equip Surface with 20% margin (m2)": total_surface_area_with_margin,
            "Swab Surface Used (m2)": swab_surface
        }]).to_excel(writer, index=False, sheet_name="Summary")
    st.download_button("Download Excel Results", output.getvalue(), "cleaning_validation_with_margin.xlsx")

    # ----------- 9. Generate Word DOCX Report -----------
    def add_heading(doc, text, level=1):
        heading = doc.add_heading(text, level=level)
        if level == 1:
            heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        return heading

    def add_section_title(doc, text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0,51,102)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        return p

    def add_table_from_df(doc, df, title=None, style="Table Grid"):
        if title:
            add_section_title(doc, title)
        table = doc.add_table(rows=1, cols=len(df.columns), style=style)
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(df.columns):
            hdr_cells[i].text = str(col)
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, val in enumerate(row):
                row_cells[i].text = str(val)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        doc.add_paragraph()
        return table

    doc = Document()
    doc.core_properties.title = "Cleaning Validation Protocol - MACO, Swab & Rinse Limit Report"

    add_heading(doc, "CLEANING VALIDATION PROTOCOL", 1)
    doc.add_paragraph("Date: .......................................")
    doc.add_paragraph()
    add_section_title(doc, "1. Introduction")
    doc.add_paragraph(
        "This report documents the systematic evaluation of cleaning validation using MACO (Maximum Allowable Carryover), Swab, and Rinse limit calculations. "
        "The protocol is based on worst-case product selection, group rating assignment for solubility, dose, toxicity, and cleaning difficulty, and includes analytical and equipment considerations. "
        "All calculations are performed according to current industry guidelines and in compliance with regulatory expectations."
    )
    add_section_title(doc, "2. Objective")
    doc.add_paragraph(
        "To establish scientifically justified cleaning limits for shared equipment, based on product and process understanding, "
        "using group ratings, worst-case identification, and three MACO calculation methods (10 ppm, TDD, ADE/PDE)."
    )
    add_section_title(doc, "3. Product Details and Group Ratings")
    add_table_from_df(
        doc,
        df[
            [
                'Product Name', 'Solubility', 'Solubility_Group',
                'Min Dose (mg)', 'Dose_Group', 'ADE/PDE (µg/day)', 'Toxicity_Group',
                'Hardest To Clean', 'Cleaning_Group',
                'Worst_Case_Rating', 'Min Batch Size (kg)', 'Max Dose (mg)', 'BatchSize_Dose_Ratio'
            ]
        ],
        title="Product Group Assignment & Rating"
    )
    add_section_title(doc, "4. Analytical Method Validation Data")
    add_table_from_df(
        doc,
        df[
            [
                'Product Name', 'Swab Recovery %', 'LOD in ppm', 'LOQ in ppm',
                'Swab Dilution in ml as per AMV', 'Swab Surface in M. Sq.'
            ]
        ],
        title="Analytical Method Parameters"
    )
    add_section_title(doc, "5. Equipment Details")
    add_table_from_df(
        doc,
        df_equip[['Eq. Name', 'Eq. ID', 'Product contact Surface Area (m2)', 'Used for', 'Cleaning   Procedure No.']],
        title="Equipment List"
    )
    add_section_title(doc, "6. Group Definitions")
    add_table_from_df(doc, solubility_template, title="Solubility Group Definition")
    add_table_from_df(doc, dose_template, title="Dose Group Definition")
    add_table_from_df(doc, toxicity_template, title="Toxicity Group Definition")
    add_table_from_df(doc, cleaning_template, title="Hardest To Clean Group Definition")
    add_section_title(doc, "7. Worst Case Product Selection")
    doc.add_paragraph("Previous Worst Case (By Highest Rating):")
    add_table_from_df(doc, prev_worst_case.to_frame().T[['Product Name', 'Worst_Case_Rating', 'Min Batch Size (kg)', 'ADE/PDE (µg/day)', 'Min Dose (mg)', 'Max Dose (mg)']])
    doc.add_paragraph("Next Worst Case (By Lowest Min Batch / Max Dose ratio):")
    add_table_from_df(doc, next_worst_case.to_frame().T[['Product Name', 'BatchSize_Dose_Ratio', 'Min Batch Size (kg)', 'ADE/PDE (µg/day)', 'Min Dose (mg)', 'Max Dose (mg)']])
    add_section_title(doc, "8. MACO (Maximum Allowable Carryover) Calculations")
    macodf = pd.DataFrame([
        ['10 ppm', f"{maco_10ppm:.4f} mg"],
        ['TDD', f"{maco_tdd:.4f} mg"],
        ['ADE/PDE', f"{maco_ade:.4f} mg"],
        ['Lowest MACO (used for limits)', f"{lowest_maco:.4f} mg"]
    ], columns=['MACO Calculation Method', 'Value (mg)'])
    add_table_from_df(doc, macodf, title="MACO Calculation Summary")
    add_section_title(doc, "9. Swab Limit Calculation")
    swabdf = pd.DataFrame([{
        "Swab Surface in M. Sq.": swab_surface,
        "Total Equip Surface with 20% margin (m2)": round(total_surface_area_with_margin, 4),
        "Lowest MACO (mg)": round(lowest_maco, 4),
        "Swab Limit (mg)": round(swab_limit, 6)
    }])
    add_table_from_df(doc, swabdf, title="Swab Limit Summary")
    add_section_title(doc, "10. Rinse Limits and Volumes per Equipment")
    add_table_from_df(doc, df_rinse_limits, title="Rinse Limits per Equipment")
    add_section_title(doc, "11. Conclusion & Summary")
    doc.add_paragraph(
        "The cleaning validation study demonstrates a robust, risk-based approach for setting carryover limits. "
        "The application of 20% margin to the total equipment surface area, group-based worst-case selection, and multiple MACO calculation approaches ensures regulatory compliance and patient safety. "
        "The lowest MACO value has been used for swab and rinse calculations. This protocol should be reviewed and approved prior to execution."
    )
    doc.add_paragraph()
    doc.add_paragraph("Prepared by:  ___________________________")
    doc.add_paragraph("Reviewed by:  __________________________")
    doc.add_paragraph("Approved by:  __________________________")

    docx_io = io.BytesIO()
    doc.save(docx_io)
    st.download_button("Download Word Report", docx_io.getvalue(), "Cleaning_Validation_Protocol_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
