import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Product Data Analysis App", layout="wide")
st.title("Product Data Analysis App")

st.markdown("""
**Upload the following files OR use the example files provided in this repo:**
- `product_details.xlsx`
- `analytical_method_validation.xlsx`
- `solubility_cleaning.xlsx`
- `equipment_details.xlsx`
- `rating_criteria.xlsx`
""")

mode = st.radio(
    "How do you want to provide data?",
    ("Use example files from repo", "Upload my own files")
)

if mode == "Upload my own files":
    product_details_file = st.file_uploader("Upload Product Details file", type=["xlsx"], key="product_details")
    amv_file = st.file_uploader("Upload Analytical Method Validation file", type=["xlsx"], key="amv")
    solubility_cleaning_file = st.file_uploader("Upload Solubility & Cleaning file", type=["xlsx"], key="sol_clean")
    equipment_details_file = st.file_uploader("Upload Equipment Details file", type=["xlsx"], key="equip")
    rating_criteria_file = st.file_uploader("Upload Rating Criteria file", type=["xlsx"], key="rating")
else:
    # Use files from repo directory
    product_details_file = "product_details.xlsx"
    amv_file = "analytical_method_validation.xlsx"
    solubility_cleaning_file = "solubility_cleaning.xlsx"
    equipment_details_file = "equipment_details.xlsx"
    rating_criteria_file = "rating_criteria.xlsx"

# Check if all files are present/selected
files_ready = (
    product_details_file and amv_file and solubility_cleaning_file and equipment_details_file and rating_criteria_file
)

if files_ready:
    # Load dataframes
    df_product = pd.read_excel(product_details_file)
    df_amv = pd.read_excel(amv_file)
    df_sol_clean = pd.read_excel(solubility_cleaning_file)
    df_equip = pd.read_excel(equipment_details_file)
    df_criteria = pd.read_excel(rating_criteria_file, sheet_name=None)

    st.success("All files loaded successfully! Here are samples from each:")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Product Details", 
        "Analytical Method Validation", 
        "Solubility & Cleaning", 
        "Equipment Details", 
        "Rating Criteria"
    ])

    with tab1:
        st.dataframe(df_product.head(20))

    with tab2:
        st.dataframe(df_amv.head(20))

    with tab3:
        st.dataframe(df_sol_clean.head(20))

    with tab4:
        st.dataframe(df_equip.head(20))

    with tab5:
        for sheet_name, sheet_df in df_criteria.items():
            st.markdown(f"### {sheet_name}")
            st.dataframe(sheet_df.head(20))
else:
    st.info("Please provide all 5 required files to proceed.")
