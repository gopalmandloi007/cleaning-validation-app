# MACO Calculation App By Gopal Mandloi

This Streamlit app provides a one-stop solution for MACO, Swab Limit, and Rinse Limit calculations in pharmaceutical cleaning validation, specifically designed for OSD (Oral Solid Dosage Form) and API facilities.

---

## 🚀 How to Use This App

### 1. Download Blank Templates
- Click on the **"Download Blank Templates"** section in the app.
- Download all provided Excel templates (`.xlsx`).
- **Do not change column names or sheet names.**

### 2. Fill Your Data
- Enter your product, equipment, and criteria data in the blank templates.
- Save your completed Excel files.

### 3. Upload Your Filled Excel Files
- In the app, select **"Upload my own files"**.
- Upload all required Excel files:
  - `product_details.xlsx`
  - `analytical_method_validation.xlsx`
  - `solubility_cleaning.xlsx`
  - `equipment_details.xlsx`
  - `rating_criteria.xlsx`

### 4. View Final Results
- Click on any result button (e.g., **Previous Worst Case Product**, **MACO Calculations**, **Swab Limit**, etc.) to see your calculation results.

---

## 📁 Required Files

| File Name                      | Content Description           |
|---------------------------------|------------------------------|
| product_details.xlsx           | Product-level data            |
| analytical_method_validation.xlsx | Analytical validation data |
| solubility_cleaning.xlsx       | Solubility & cleaning data    |
| equipment_details.xlsx         | Equipment surface area etc.   |
| rating_criteria.xlsx           | 4 sheets: Solubility, Dose, Toxicity, Cleaning |

> Use only the provided blank templates for data entry to avoid errors.

---

## 📝 Common Issues

- **Column or Sheet Name Error?**  
  Use the blank templates and verify names carefully.
- **File Format:**  
  All files must be `.xlsx` (Excel) format.
- **Calculation Check:**  
  Always verify results manually for critical use.

---

## ⚠️ Disclaimer

> **This app is under development and may show wrong calculations.  
> Always calculate manually and match both the results.  
> I am not responsible for any incorrect calculation.  
> This app is developed for OSD (Oral Solid Dosage Form & API) only, not for other formulations.**

If you find/observe any error, please inform:
- **Gopal Mandloi**  
  WhatsApp/Mobile: **9827276040**

---

## 💻 Deployment

1. Clone this repository:
    ```bash
    git clone https://github.com/gopalmandloi007/maco-calculation-app.git
    cd maco-calculation-app
    ```
2. Install requirements:
    ```bash
    pip install -r requirements.txt
    ```
3. Run the app:
    ```bash
    streamlit run app.py
    ```

---

## 🌐 Contact

For suggestions, custom app development, or reporting issues, contact **Gopal Mandloi** (WhatsApp: 9827276040).

---
