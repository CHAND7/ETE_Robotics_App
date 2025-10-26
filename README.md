# ETE Robotics RFQ Management App

This is an internal Streamlit web application for ETE Robotics to manage RFQs (Request for Quotation) efficiently.

### Features:
- User & Admin Login System  
- Excel-driven dropdown data (BOM, Component categories, etc.)  
- Multi-step RFQ form (Step 1–3 with breadcrumb navigation)  
- Automatic PDF and PowerPoint generation with ETE Robotics branding  
- Auto email dispatch of RFQs  
- Data stored securely and locally

### Files Included:
- `app.py` – Main Streamlit application  
- `ETE_Robotics-Bom-Data-for-softwares-development.xlsx` – BOM data source  
- `ETE-RFQ_CheckList.xlsx` – Base checklist for RFQ form  
- `ETE-Robotics-Logo.png` – Branding logo for reports  
- `requirements.txt` – Dependencies for Streamlit deployment  

### Deployment:
The app is deployable directly via [Streamlit Cloud](https://streamlit.io/cloud).

### Local Development:
```bash
pip install -r requirements.txt
streamlit run app.py
