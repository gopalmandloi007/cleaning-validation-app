import streamlit as st
import pandas as pd

st.title("Product Data Analysis App")

# File uploaders for all required files
product_details_file = st.file_uploader("Upload Product Details file", type=["xlsx"])
amv_file = st.file_uploader("Upload Analytical Method Validation file", type=["xlsx"])
solubility_cleaning_file = st.file_uploader("Upload Solubility & Cleaning file", type=["xlsx"])
equipment_details_file = st.file_uploader("Upload Equipment Details file", type=["xlsx"])
rating_criteria_file = st.file_uploader("Upload Rating Criteria file", type=["xlsx"])

if all([product_details_file, amv_file, solubility_cleaning_file, equipment_details_file, rating_criteria_file]):
    # Read each file into a DataFrame
    df_product = pd.read_excel(product_details_file)
    df_amv = pd.read_excel(amv_file)
    df_sol_clean = pd.read_excel(solubility_cleaning_file)
    df_equip = pd.read_excel(equipment_details_file)
    # For rating_criteria, read all sheets into a dict of DataFrames
    df_criteria = pd.read_excel(rating_criteria_file, sheet_name=None)

    st.success("All files uploaded successfully! Here are samples from each:")

    st.subheader("Product Details")
    st.dataframe(df_product.head())

    st.subheader("Analytical Method Validation")
    st.dataframe(df_amv.head())

    st.subheader("Solubility & Cleaning")
    st.dataframe(df_sol_clean.head())

    st.subheader("Equipment Details")
    st.dataframe(df_equip.head())

    st.subheader("Rating Criteria (All Sheets)")
    for sheet_name, sheet_df in df_criteria.items():
        st.markdown(f"**{sheet_name}**")
        st.dataframe(sheet_df.head())

else:
    st.info("Please upload all 5 required files to proceed.")
