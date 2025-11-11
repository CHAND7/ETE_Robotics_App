# app.py
import streamlit as st
import pandas as pd
import io
import os
import re
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

# ----------------- CONFIG / SECRETS -----------------
ADMIN_USERNAME = st.secrets.get("admin", {}).get("username", "admin")
ADMIN_PASSWORD = st.secrets.get("admin", {}).get("password", "ete123")

EXCEL_FILE = "ETE_Robotics-Bom-Data-for-softwares-development.xlsx"
LOGO_FILE = "ETE-Robotics-Logo.png"
PPT_TEMPLATE = "ETE_Robotics_Proposal_Customer-Name_Date-Revision.pptx"  # optional template

# ----------------- UTIL: Load BOM Robustly -----------------
@st.cache_data
def load_excel_data(path=EXCEL_FILE):
    if not os.path.exists(path):
        return pd.DataFrame()
    for header_idx in (11, 12, 10, 9):
        try:
            df = pd.read_excel(path, header=header_idx, usecols="B:H", engine="openpyxl")
            df.columns = [str(c).strip().replace("\n", " ").replace("\r", "") for c in df.columns]
            rename_map = {}
            for col in df.columns:
                low = col.lower()
                if "head" in low:
                    rename_map[col] = "Head"
                elif "description" in low:
                    rename_map[col] = "Description"
                elif "model" in low or "key spec" in low:
                    rename_map[col] = "Model/Key Spec"
                elif "unit" in low and "cost" in low:
                    rename_map[col] = "Unit Cost"
                elif "qty" == low or "quantity" in low:
                    rename_map[col] = "Qty"
                elif "s.no" in low or "sno" in low:
                    rename_map[col] = "S.no"
            df = df.rename(columns=rename_map)
            if "Head" in df.columns and "Model/Key Spec" in df.columns:
                df = df.dropna(subset=["Head"])
                if "Unit Cost" in df.columns:
                    df["Unit Cost"] = pd.to_numeric(df["Unit Cost"].astype(str).str.replace(r'[^\d\.\-]', '', regex=True), errors="coerce").fillna(0)
                return df
        except Exception:
            continue
    # fallback attempt
    try:
        df2 = pd.read_excel(path, header=12, engine="openpyxl")
        df2.columns = [str(c).strip().replace("\n", " ").replace("\r", "") for c in df2.columns]
        return df2
    except Exception:
        return pd.DataFrame()

bom_df = load_excel_data()

# ----------------- Helpers -----------------
def split_spec_values(cell):
    if not cell or pd.isna(cell):
        return []
    cell = str(cell)
    parts = re.split(r'\s*\|\s*|\s+I\s+|/|;|,', cell)
    return [p.strip() for p in parts if p.strip()]

@st.cache_data
def build_model_options(df):
    options = []
    if df is None or df.empty: 
        return options
    col = "Model/Key Spec"
    if col not in df.columns:
        return options
    for val in df[col].dropna().astype(str):
        options.extend(split_spec_values(val))
    return sorted(list(set(options)))

MODEL_OPTIONS = build_model_options(bom_df)

def find_unit_cost_for_model(df, chosen_model):
    if df is None or df.empty or not chosen_model:
        return 0.0
    mask = df["Model/Key Spec"].astype(str).str.contains(re.escape(chosen_model), case=False, na=False)
    if mask.any():
        vals = df.loc[mask, "Unit Cost"].dropna().astype(float)
        if not vals.empty:
            return float(vals.iloc[0])
    return 0.0

# ----------------- UI helpers -----------------
BREADCRUMB_CSS = """
<style>
.bc-row { display:flex; justify-content:center; gap:10px; margin-top:10px; margin-bottom:6px; }
.bc-btn { padding:6px 12px; border-radius:8px; background:#f1f3f4; cursor:pointer; font-weight:600; }
.bc-active { background:linear-gradient(90deg,#ff8800,#ffb06b); color:white; box-shadow:0 2px 6px rgba(0,0,0,0.12);}
.logo-small { padding: 6px 0; }
</style>
"""

# ----------------- PDF generator (defensive) -----------------
def ensure_table_data(rows):
    # Ensure at least one row and one column: convert empty dict to [[ "No data", "" ]]
    if not rows or (isinstance(rows, list) and len(rows) == 0):
        return [["", ""]]
    return rows

def create_pdf(customer_info, requirements, rfq_checklist, items):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []
    # logo
    if os.path.exists(LOGO_FILE):
        try:
            story.append(RLImage(LOGO_FILE, width=150, height=45))
        except Exception:
            pass
    story.append(Spacer(1, 6))
    story.append(Paragraph("ETE Robotics — RFQ Summary", styles['Title']))
    story.append(Spacer(1, 10))

    # Customer Info
    cust_rows = []
    if customer_info:
        for k, v in customer_info.items():
            cust_rows.append([Paragraph(f"<b>{k}</b>", styles['Normal']), Paragraph(str(v), styles['Normal'])])
    cust_rows = ensure_table_data(cust_rows)
    t = Table(cust_rows, colWidths=[120, 360])
    t.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('INNERGRID', (0,0), (-1,-1), 0.25, colors.grey)]))
    story.append(t)
    story.append(Spacer(1, 8))

    # Requirements
    req_rows = []
    if requirements:
        for k, v in requirements.items():
            req_rows.append([Paragraph(f"{k}", styles['Normal']), Paragraph(str(v), styles['Normal'])])
    req_rows = ensure_table_data(req_rows)
    story.append(Paragraph("<b>Requirements</b>", styles['Heading3']))
    rt = Table(req_rows, colWidths=[160, 320])
    rt.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
    story.append(rt)
    story.append(Spacer(1, 8))

    # RFQ Checklist
    chk_rows = []
    if rfq_checklist:
        for k, v in rfq_checklist.items():
            chk_rows.append([Paragraph(str(k), styles['Normal']), Paragraph(str(v), styles['Normal'])])
    chk_rows = ensure_table_data(chk_rows)
    story.append(Paragraph("<b>RFQ Checklist</b>", styles['Heading3']))
    ct = Table(chk_rows, colWidths=[160, 320])
    ct.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
    story.append(ct)
    story.append(Spacer(1, 8))

    # Items table
    if items:
        table_data = [["S.no", "Head", "Model/Spec", "Qty", "Unit Cost", "Line Cost (INR)"]]
        total = 0.0
        for it in items:
            table_data.append([it.get("S.no",""), it.get("Head",""), it.get("ModelSpec",""), str(it.get("Qty","")), f"{it.get('UnitCost',0):,.2f}", f"{it.get('LineCost',0):,.2f}"])
            total += float(it.get("LineCost", 0))
        table_data.append(["", "", "", "", "Total", f"{total:,.2f}"])
        tbl = Table(table_data, colWidths=[40, 140, 160, 50, 80, 90])
        tbl.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.3,colors.grey), ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#f5f5f5")), ('ALIGN',(-2,1),(-1,-1),'RIGHT')]))
        story.append(tbl)
        story.append(Spacer(1, 8))

    doc.build(story)
    buf.seek(0)
    return buf

# ----------------- PPT generator (uses template if present) -----------------
def create_ppt(customer_info, requirements, rfq_checklist, items):
    prs = None
    if os.path.exists(PPT_TEMPLATE):
        try:
            prs = Presentation(PPT_TEMPLATE)
        except Exception:
            prs = Presentation()
    else:
        prs = Presentation()

    # Title slide (use layout 0 if available)
    try:
        slide_layout = prs.slide_layouts[0]
    except Exception:
        slide_layout = prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    try:
        slide.shapes.title.text = "RFQ / Proposal"
    except Exception:
        pass
    # Subtitle or details
    try:
        if slide.placeholders:
            for ph in slide.placeholders:
                # set the first placeholder text if empty
                try:
                    ph.text = f"{customer_info.get('Customer Name','')}"
                    break
                except Exception:
                    continue
    except Exception:
        pass
    # Add content slide for customer info
    slide = prs.slides.add_slide(prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[1])
    slide.shapes.title.text = "Customer & Requirements"
    tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(4)).text_frame
    for k,v in customer_info.items():
        p = tx.add_paragraph(); p.text = f"{k}: {v}"; p.level = 0
    tx.add_paragraph(); tx.add_paragraph()
    p = tx.add_paragraph(); p.text = "Requirements:"; p.level = 0
    for k,v in requirements.items():
        pp = tx.add_paragraph(); pp.text = f"{k}: {v}"; pp.level = 1

    # Items slide
    if items:
        slide = prs.slides.add_slide(prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[1])
        slide.shapes.title.text = "Selected Items & Budget"
        rows = len(items) + 2
        cols = 5
        left = Inches(0.4)
        top = Inches(1.4)
        width = Inches(9)
        height = Inches(0.8)
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        table.cell(0,0).text = "S.no"
        table.cell(0,1).text = "Model/Spec"
        table.cell(0,2).text = "Head"
        table.cell(0,3).text = "Qty"
        table.cell(0,4).text = "Line Cost (INR)"
        r = 1; total = 0.0
        for it in items:
            table.cell(r,0).text = str(it.get("S.no",""))
            table.cell(r,1).text = str(it.get("ModelSpec",""))
            table.cell(r,2).text = str(it.get("Head",""))
            table.cell(r,3).text = str(it.get("Qty",""))
            table.cell(r,4).text = f"{it.get('LineCost',0):,.2f}"
            total += float(it.get('LineCost',0))
            r += 1
        # last row total
        if r <= rows-1:
            table.cell(r,2).text = "Total"
            table.cell(r,4).text = f"{total:,.2f}"

    # Checklist slide
    slide = prs.slides.add_slide(prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[1])
    slide.shapes.title.text = "RFQ Checklist"
    tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(4)).text_frame
    for k,v in rfq_checklist.items():
        p = tx.add_paragraph(); p.text = f"{k}: {v}"

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# ----------------- Session state -----------------
if "step" not in st.session_state: st.session_state.step = 1
if "customer_info" not in st.session_state: st.session_state.customer_info = {}
if "requirements" not in st.session_state: st.session_state.requirements = {}
if "rfq_checklist" not in st.session_state: st.session_state.rfq_checklist = {}
if "selected_items" not in st.session_state: st.session_state.selected_items = []

# ----------------- Top layout: logo (small left) + breadcrumb center -----------------
st.set_page_config(page_title="ETE RFQ Builder", layout="wide")
top1, top2 = st.columns([1, 6])
with top1:
    if os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=110)
with top2:
    st.markdown(BREADCRUMB_CSS, unsafe_allow_html=True)
    cols = st.columns([1,1,1])
    labels = ["Step 1: Customer Info", "Step 2: RFQ Checklist", "Step 3: Submit & Generate"]
    for idx, c in enumerate(cols, start=1):
        with c:
            if st.session_state.step == idx:
                if st.button(labels[idx-1], key=f"bc{idx}"):
                    st.session_state.step = idx
            else:
                if st.button(labels[idx-1], key=f"bc{idx}_off"):
                    st.session_state.step = idx

st.write("---")

# ----------------- Step 1 -----------------
if st.session_state.step == 1:
    st.header("Step 1 — Customer Info & Requirements")
    with st.form("step1"):
        c1, c2 = st.columns(2)
        with c1:
            rfq_ref = st.text_input("RFQ Reference", value=st.session_state.customer_info.get("RFQ Reference", f"RFQ/ETE/{datetime.now().year}-{datetime.now().strftime('%m%d%H%M')}"))
            cust_name = st.text_input("Customer Name", value=st.session_state.customer_info.get("Customer Name",""))
            contact_no = st.text_input("Contact No.", value=st.session_state.customer_info.get("Contact No.",""))
            email = st.text_input("Email ID", value=st.session_state.customer_info.get("Email ID",""))
            location = st.text_input("Location", value=st.session_state.customer_info.get("Location",""))
        with c2:
            date_val = st.date_input("Date", value=st.session_state.customer_info.get("Date", datetime.today()))
            application = st.selectbox("Application", ["Robotic","SPM","Testing","Conveyor","Plant facility Transfer","Modification","Service","Std item Supply"])
            eq_type = st.selectbox("Type of Equipment", ["Hydraulic","Pneumatic","Servo","Other","Conveyor (Belt)","Conveyor (Slat)","Conveyor (Roller)"])
            product_info = st.text_input("Product Information", value=st.session_state.customer_info.get("Product Information",""))
            new_mod = st.selectbox("New / Modification", ["New","Modification"])
        col1, col2, col3 = st.columns([1,1,1])
        with col1:
            save = st.form_submit_button("Save")
        with col2:
            save_next = st.form_submit_button("Save & Next")
        with col3:
            reset = st.form_submit_button("Reset")
    if save or save_next:
        st.session_state.customer_info = {
            "RFQ Reference": rfq_ref,
            "Customer Name": cust_name,
            "Contact No.": contact_no,
            "Email ID": email,
            "Date": str(date_val),
            "Location": location,
            "Application": application,
            "Type of Equipment": eq_type,
            "Product Information": product_info,
            "New/Modification": new_mod
        }
        st.success("Step 1 data saved.")
        if save_next:
            st.session_state.step = 2
            try:
                st.experimental_rerun()
            except Exception:
                pass
    if reset:
        st.session_state.customer_info = {}
        try:
            st.experimental_rerun()
        except Exception:
            pass

# ----------------- Step 2 -----------------
elif st.session_state.step == 2:
    st.header("Step 2 — RFQ Checklist & Bill of Quantity")
    with st.form("step2"):
        st.subheader("A) General Information")
        g1, g2 = st.columns(2)
        with g1:
            project_desc = st.text_input("Project Description", value=st.session_state.rfq_checklist.get("Project Description",""))
            proposal_no = st.text_input("Proposal No.", value=st.session_state.rfq_checklist.get("Proposal No","P-001"))
            assigned_to = st.text_input("Assigned To", value=st.session_state.rfq_checklist.get("Assigned To",""))
        with g2:
            customer_repeat = st.text_input("Customer (repeat)", value=st.session_state.rfq_checklist.get("Customer", st.session_state.customer_info.get("Customer Name","")))
            chk_date = st.date_input("Date", value=datetime.today())

        st.subheader("B) Project Summary")
        s1, s2 = st.columns(2)
        with s1:
            ps_type = st.selectbox("Type of Equipment (summary)", ["Hydraulic","Pneumatic","Servo","Other"])
            ps_application = st.selectbox("Application (summary)", ["Robotic","SPM","Testing","Conveyor"])
            cycle_time = st.text_input("Cycle Time", value=st.session_state.rfq_checklist.get("Cycle Time",""))
        with s2:
            no_of_variants = st.number_input("No. of Variants", min_value=1, value=int(st.session_state.rfq_checklist.get("No. of Variants",1)))
            delivery_time = st.text_input("Delivery Time", value=st.session_state.rfq_checklist.get("Delivery Time",""))
            i_and_c_location = st.text_input("I&C Location", value=st.session_state.rfq_checklist.get("I&C Location",""))

        st.subheader("C) Concept Layout")
        concept_layout = st.text_area("Concept Layout (notes)", value=st.session_state.rfq_checklist.get("Concept Layout",""))

        st.subheader("D) Process Details / Key Features")
        key_features = st.text_area("Key Features / Process Details", value=st.session_state.rfq_checklist.get("Key Features",""))

        st.subheader("E) Bill of Quantity")
        b1, b2, b3 = st.columns([2,1,1])
        with b1:
            chosen_model = st.selectbox("Select Model / Key Spec (from BOM)", ["-- select --"] + MODEL_OPTIONS, index=0)
        with b2:
            chosen_qty = st.number_input("Qty", min_value=1, value=1)
        with b3:
            add_btn = st.form_submit_button("Add Item to BOM")

        st.markdown("**Current selected items**")
        if st.session_state.selected_items:
            st.dataframe(pd.DataFrame(st.session_state.selected_items)[["ModelSpec","Head","Qty","UnitCost","LineCost"]], use_container_width=True)
        else:
            st.info("No BOM items added yet. Use the selector above to add items from BOM.")

        c1, c2, c3 = st.columns([1,1,1])
        with c1:
            save2 = st.form_submit_button("Save")
        with c2:
            save_next2 = st.form_submit_button("Save & Next")
        with c3:
            back = st.form_submit_button("Back to Step 1")

    if add_btn and chosen_model and chosen_model != "-- select --":
        unit_cost = find_unit_cost_for_model(bom_df, chosen_model)
        line_cost = unit_cost * chosen_qty
        head_val = ""
        if not bom_df.empty:
            mask = bom_df["Model/Key Spec"].astype(str).str.contains(re.escape(chosen_model), case=False, na=False)
            if mask.any():
                try:
                    head_val = bom_df.loc[mask, "Head"].iloc[0]
                except Exception:
                    head_val = ""
        item = {"S.no": len(st.session_state.selected_items)+1, "ModelSpec": chosen_model, "Head": head_val, "Qty": int(chosen_qty), "UnitCost": float(unit_cost), "LineCost": float(line_cost)}
        st.session_state.selected_items.append(item)
        st.success(f"Added {chosen_model} x {chosen_qty}")

    if save2:
        st.session_state.rfq_checklist = {
            "Project Description": project_desc,
            "Proposal No": proposal_no,
            "Assigned To": assigned_to,
            "Customer": customer_repeat,
            "Date": str(chk_date),
            "Project Summary": {"Type": ps_type, "Application": ps_application, "Cycle Time": cycle_time, "Variants": no_of_variants, "Delivery Time": delivery_time, "I&C Location": i_and_c_location},
            "Concept Layout": concept_layout,
            "Key Features": key_features
        }
        st.success("Step 2 data saved.")
    if save_next2:
        st.session_state.rfq_checklist = {
            "Project Description": project_desc,
            "Proposal No": proposal_no,
            "Assigned To": assigned_to,
            "Customer": customer_repeat,
            "Date": str(chk_date),
            "Project Summary": {"Type": ps_type, "Application": ps_application, "Cycle Time": cycle_time, "Variants": no_of_variants, "Delivery Time": delivery_time, "I&C Location": i_and_c_location},
            "Concept Layout": concept_layout,
            "Key Features": key_features
        }
        st.session_state.step = 3
        try:
            st.experimental_rerun()
        except Exception:
            pass
    if back:
        st.session_state.step = 1
        try:
            st.experimental_rerun()
        except Exception:
            pass

# ----------------- Step 3 -----------------
elif st.session_state.step == 3:
    st.header("Step 3 — Review & Generate (No Email)")
    st.subheader("Review Entered Data")
    st.markdown("**Customer Info**"); st.json(st.session_state.customer_info)
    st.markdown("**Requirements**"); st.json(st.session_state.requirements)
    st.markdown("**RFQ Checklist**"); st.json(st.session_state.rfq_checklist)
    st.markdown("**Selected BOM Items**")
    if st.session_state.selected_items:
        st.dataframe(pd.DataFrame(st.session_state.selected_items)[["ModelSpec","Head","Qty","UnitCost","LineCost"]], use_container_width=True)
    else:
        st.info("No items selected.")

    total = sum([it.get("LineCost",0) for it in st.session_state.selected_items]) if st.session_state.selected_items else 0.0
    st.markdown(f"### Total Estimated Cost: **₹ {total:,.2f}**")

    c1, c2 = st.columns([1,1])
    with c1:
        if st.button("Back to Step 2"):
            st.session_state.step = 2
            try:
                st.experimental_rerun()
            except Exception:
                pass
    with c2:
        if st.button("Generate PDF & PPT"):
            pdf_bytes = create_pdf(st.session_state.customer_info, st.session_state.requirements, st.session_state.rfq_checklist, st.session_state.selected_items)
            ppt_bytes = create_ppt(st.session_state.customer_info, st.session_state.requirements, st.session_state.rfq_checklist, st.session_state.selected_items)
            st.success("Generated PDF & PPT. Download below.")
            st.download_button("Download PDF", data=pdf_bytes, file_name=f"{st.session_state.customer_info.get('RFQ Reference','RFQ')}.pdf", mime="application/pdf")
            st.download_button("Download PPTX", data=ppt_bytes, file_name=f"{st.session_state.customer_info.get('RFQ Reference','RFQ')}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            st.session_state.last_pdf = pdf_bytes.getvalue()
            st.session_state.last_pptx = ppt_bytes.getvalue()
