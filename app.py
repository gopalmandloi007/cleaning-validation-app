import streamlit as st
import pandas as pd

st.set_page_config(page_title="MACO Calculation App By Gopal Mandloi", layout="wide")
st.markdown("<h1 style='text-align: center; color: #2E8B57;'>MACO Calculation App By Gopal Mandloi</h1>", unsafe_allow_html=True)

st.markdown("""
Upload your files or use example files:

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
        st.error(f"Error loading file: {e}")
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
    templates = None
    try:
        templates = pd.read_excel(uploaded_criteria, sheet_name=None)
    except Exception as e:
        st.error(f"Error loading Rating Criteria: {e}")

    if any(x is None for x in [df, df_equip, templates]):
        st.stop()

    solubility_template = templates['Solubility']
    dose_template = templates['Dose']
    toxicity_template = templates['Toxicity']
    cleaning_template = templates['Cleaning']

    # Assign groups
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

    # Calculations
    prev_worst_case = df.loc[df['Worst_Case_Rating'].idxmax()]
    next_worst_case = df.loc[df['BatchSize_Dose_Ratio'].idxmin()]
    min_batch_next_kg = next_worst_case['Min Batch Size (kg)']
    max_dose_next_mg = next_worst_case['Max Dose (mg)']
    min_dose_prev_mg = prev_worst_case['Min Dose (mg)']
    ade_prev_ug = prev_worst_case['ADE/PDE (µg/day)']
    ade_prev_mg = ade_prev_ug / 1000

    maco_10ppm = 0.00001 * min_batch_next_kg * 1e6 / max_dose_next_mg
    maco_tdd = min_dose_prev_mg * min_batch_next_kg * 1e6 / (max_dose_next_mg * 1000)
    maco_ade = ade_prev_mg * min_batch_next_kg * 1e6 / max_dose_next_mg
    lowest_maco = min(maco_10ppm, maco_tdd, maco_ade)

    total_surface_area = df_equip['Product contact Surface Area (m2)'].sum()
    total_surface_area_with_margin = total_surface_area * 1.2
    swab_surface = df['Swab Surface in M. Sq.'].iloc[0]
    swab_limit = lowest_maco * swab_surface / total_surface_area_with_margin

    # Rinse limit & volume per equipment
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

    st.markdown("---")
    st.subheader("View Final Results")

    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        if st.button("📦 Previous Worst Case Product", use_container_width=True):
            st.success(f"**Previous Worst Case Product**: {prev_worst_case['Product Name']}")
            st.write(f"Min Dose: {prev_worst_case['Min Dose (mg)']} mg")
            st.write(f"Max Dose: {prev_worst_case['Max Dose (mg)']} mg")
            st.write(f"ADE/PDE: {prev_worst_case['ADE/PDE (µg/day)']} µg/day")
    with col2:
        if st.button("🚚 Next Worst Case Product", use_container_width=True):
            st.success(f"**Next Worst Case Product**: {next_worst_case['Product Name']}")
            st.write(f"Min Batch Size: {next_worst_case['Min Batch Size (kg)']} kg")
            st.write(f"Max Dose: {next_worst_case['Max Dose (mg)']} mg")
    with col3:
        if st.button("🛑 MACO Calculations", use_container_width=True):
            st.success("**MACO Results**")
            st.write(f"10 ppm MACO: {maco_10ppm:.4f} mg")
            st.write(f"TDD MACO: {maco_tdd:.4f} mg")
            st.write(f"ADE/PDE MACO: {maco_ade:.4f} mg")
            st.write(f"**Lowest MACO (used): {lowest_maco:.4f} mg**")

    col4, col5 = st.columns([1, 1])
    with col4:
        if st.button("🧪 Swab Limit", use_container_width=True):
            st.success("**Swab Limit**")
            st.write(f"Swab Surface Used: {swab_surface} m²")
            st.write(f"Total Equip Surface (with 20% margin): {total_surface_area_with_margin:.2f} m²")
            st.write(f"**Swab Limit: {swab_limit:.6f} mg**")
    with col5:
        if st.button("💧 Rinse Limit & Volume (Equipment-wise)", use_container_width=True):
            st.success("**Rinse Limit & Volume per Equipment**")
            st.dataframe(df_rinse_limits, use_container_width=True)

else:
    st.info("Please upload all 5 required files to proceed.")
