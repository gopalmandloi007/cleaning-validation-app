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

def safe_read_excel(path_or_buffer, **kwargs):
    try:
        df = pd.read_excel(path_or_buffer, **kwargs)
        return df
    except FileNotFoundError:
        st.error(f"File not found: `{path_or_buffer}`. Please make sure the file exists in your repo or upload it.")
        return None
    except ValueError as ve:
        st.error(f"ValueError while reading `{path_or_buffer}`: {ve}")
        return None
    except Exception as e:
        st.error(f"Error reading `{path_or_buffer}`: {e}")
        return None

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
    # Use files from repo directory, but check if they exist!
    def check_file(name):
        if not os.path.exists(name):
            st.error(f"File `{name}` is missing from your repo! Please add it or upload your own files.")
            return None
        return name

    product_details_file = check_file("product_details.xlsx")
    amv_file = check_file("analytical_method_validation.xlsx")
    solubility_cleaning_file = check_file("solubility_cleaning.xlsx")
    equipment_details_file = check_file("equipment_details.xlsx")
    rating_criteria_file = check_file("rating_criteria.xlsx")

files_ready = all([
    product_details_file, amv_file, solubility_cleaning_file, equipment_details_file, rating_criteria_file
])

if files_ready:
    # Try to read each file, handle all errors gracefully
    df_product = safe_read_excel(product_details_file)
    df_amv = safe_read_excel(amv_file)
    df_sol_clean = safe_read_excel(solubility_cleaning_file)
    df_equip = safe_read_excel(equipment_details_file)
    df_criteria = safe_read_excel(rating_criteria_file, sheet_name=None)

    # Check that all DataFrames loaded successfully
    if None in [df_product, df_amv, df_sol_clean, df_equip, df_criteria]:
        st.stop()

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
