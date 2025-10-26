# app.py
import streamlit as st
import pandas as pd
import os
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image

# ------------------- CONFIGURATION -------------------
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "ete123"

# Load Excel data
EXCEL_FILE = "ETE_Robotics-Bom-Data-for-softwares-development.xlsx"
LOGO_FILE = "ETE-Robotics-Logo.png"

# Load Excel data once
@st.cache_data
def load_excel_data():
    df = pd.read_excel(EXCEL_FILE, header=12, usecols="B:H")
    df = df.rename(columns={
        'Unnamed: 1': 'Head',
        'Unnamed: 2': 'Description',
        'Unnamed: 4': 'Model/Key Spec',
        'Unnamed: 6': 'Unit Cost'
    })
    df = df.dropna(subset=['Head'])
    return df

data = load_excel_data()

# ------------------- HELPER FUNCTIONS -------------------
def get_dropdown_options(df, head_name):
    sub_df = df[df['Head'].str.contains(head_name, case=False, na=False)]
    options = []
    for val in sub_df['Model/Key Spec'].dropna():
        for opt in str(val).split('|'):
            options.append(opt.strip())
    return sorted(list(set(options)))

def generate_pdf(customer_info, requirements, rfq_checklist, total_cost):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Logo
    if os.path.exists(LOGO_FILE):
        c.drawImage(LOGO_FILE, 50, height - 100, width=150, preserveAspectRatio=True, mask='auto')

    c.setFont("Helvetica-Bold", 16)
    c.drawString(220, height - 80, "RFQ Summary Report")

    c.setFont("Helvetica", 10)
    y = height - 140
    c.drawString(50, y, "Customer Details:")
    y -= 20
    for k, v in customer_info.items():
        c.drawString(60, y, f"{k}: {v}")
        y -= 15

    y -= 10
    c.drawString(50, y, "Requirements:")
    y -= 20
    for k, v in requirements.items():
        c.drawString(60, y, f"{k}: {v}")
        y -= 15

    y -= 10
    c.drawString(50, y, "RFQ Checklist:")
    y -= 20
    for k, v in rfq_checklist.items():
        c.drawString(60, y, f"{k}: {v}")
        y -= 15

    y -= 30
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, f"Total Estimated Cost: ₹ {total_cost:,.2f}")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

def generate_ppt(customer_info, requirements, rfq_checklist, total_cost):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title = slide.shapes.title
    title.text = "RFQ Summary - ETE Robotics"

    left, top = Inches(1), Inches(1.5)
    txBox = slide.shapes.add_textbox(left, top, Inches(8), Inches(5))
    tf = txBox.text_frame
    tf.word_wrap = True

    tf.text = "Customer Information:\n"
    for k, v in customer_info.items():
        p = tf.add_paragraph()
        p.text = f"{k}: {v}"
        p.level = 1

    tf.add_paragraph()
    p = tf.add_paragraph("Requirements:")
    for k, v in requirements.items():
        sub = tf.add_paragraph()
        sub.text = f"{k}: {v}"
        sub.level = 1

    tf.add_paragraph()
    p = tf.add_paragraph("RFQ Checklist:")
    for k, v in rfq_checklist.items():
        sub = tf.add_paragraph()
        sub.text = f"{k}: {v}"
        sub.level = 1

    tf.add_paragraph()
    p = tf.add_paragraph(f"Total Estimated Cost: ₹ {total_cost:,.2f}")
    p.font.bold = True

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

def send_email(receiver_email, subject, body, attachments):
    sender_email = st.secrets["email"]["sender"]
    password = st.secrets["email"]["password"]
    smtp_server = st.secrets["email"]["smtp_server"]
    smtp_port = st.secrets["email"]["smtp_port"]

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    for filename, filedata in attachments.items():
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(filedata.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, password)
    server.send_message(msg)
    server.quit()

# ------------------- AUTHENTICATION -------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

st.title("🔧 ETE Robotic Systems Integrator - RFQ Automation Platform")

if not st.session_state.logged_in:
    st.subheader("Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            st.session_state.logged_in = True
            st.session_state.user_role = "admin"
        elif username == "user" and password == "user123":
            st.session_state.logged_in = True
            st.session_state.user_role = "user"
        else:
            st.error("Invalid credentials")
    st.stop()

# ------------------- BREADCRUMB NAVIGATION -------------------
page = st.sidebar.radio("Navigation", ["Step 1: Customer Info & Requirements", "Step 2: RFQ Checklist", "Step 3: Submit & Generate"])

# ------------------- STEP 1 -------------------
if page.startswith("Step 1"):
    st.header("Step 1: Customer Information & Requirements")

    customer_info = {
        "RFQ Reference": st.text_input("RFQ Reference (Auto-generated)", f"RFQ/ETE/{datetime.now().year}-{datetime.now().strftime('%m%d%H%M')}"),
        "Customer Name": st.text_input("Customer Name"),
        "Contact No.": st.text_input("Contact Number"),
        "Email ID": st.text_input("Email ID"),
        "Date": st.date_input("Date", datetime.today()),
        "Location": st.text_input("Location")
    }

    st.subheader("Requirements")
    requirements = {
        "Application": st.selectbox("Application", ["Robotic", "SPM", "Testing", "Conveyor", "Plant Facility Transfer", "Modification", "Service", "Std Item Supply"]),
        "Type of Equipment": st.selectbox("Type of Equipment", ["Hydraulic", "Pneumatic", "Servo", "Other", "Conveyor (Belt/Slat/Roller)"]),
        "Product Information": st.text_input("Product Information"),
        "New / Modification": st.text_input("New / Modification"),
        "Cycle Time": st.text_input("Cycle Time"),
        "Delivery Time": st.text_input("Delivery Time"),
        "Installation Location": st.text_input("Installation Location"),
        "Domestic / Export": st.selectbox("Domestic / Export", ["Domestic", "Export"]),
        "Payment Terms": st.selectbox("Payment Terms", ["30% Advance", "50% Advance", "100% on Delivery"]),
        "Transportation Scope": st.selectbox("Transportation Scope", ["Included", "Excluded"])
    }

    st.session_state.customer_info = customer_info
    st.session_state.requirements = requirements
    st.success("Step 1 saved successfully. Proceed to Step 2 from sidebar.")

# ------------------- STEP 2 -------------------
elif page.startswith("Step 2"):
    st.header("Step 2: RFQ Checklist")

    rfq_checklist = {
        "Project Description": st.text_input("Project Description"),
        "Proposal No.": st.selectbox("Proposal No.", ["P-001", "P-002", "P-003"]),
        "Assigned To": st.text_input("Assigned To"),
        "Customer": st.text_input("Customer"),
        "Date": st.date_input("Date", datetime.today())
    }

    mech_options = get_dropdown_options(data, "Mechanical")
    selected_mech = st.selectbox("Mechanical", mech_options)

    qty = st.number_input("Quantity", min_value=1, value=1)
    unit_cost = float(data[data["Model/Key Spec"].str.contains(selected_mech, na=False)]["Unit Cost"].fillna(0).iloc[0])
    total_cost = unit_cost * qty

    rfq_checklist["Mechanical Selection"] = selected_mech
    rfq_checklist["Unit Cost"] = unit_cost
    rfq_checklist["Quantity"] = qty

    st.session_state.rfq_checklist = rfq_checklist
    st.session_state.total_cost = total_cost

    st.success(f"Step 2 saved successfully. Estimated Cost: ₹ {total_cost:,.2f}. Proceed to Step 3.")

# ------------------- STEP 3 -------------------
else:
    st.header("Step 3: Submit, Generate & Email")

    if "customer_info" not in st.session_state:
        st.error("Please complete Step 1 first.")
        st.stop()
    if "rfq_checklist" not in st.session_state:
        st.error("Please complete Step 2 first.")
        st.stop()

    customer_info = st.session_state.customer_info
    requirements = st.session_state.requirements
    rfq_checklist = st.session_state.rfq_checklist
    total_cost = st.session_state.total_cost

    pdf_buffer = generate_pdf(customer_info, requirements, rfq_checklist, total_cost)
    ppt_buffer = generate_ppt(customer_info, requirements, rfq_checklist, total_cost)

    st.download_button("📄 Download PDF", data=pdf_buffer, file_name="RFQ_Summary.pdf")
    st.download_button("📊 Download PPT", data=ppt_buffer, file_name="RFQ_Summary.pptx")

    receiver = st.text_input("Send Email To:")
    if st.button("Send Email"):
        send_email(receiver, "RFQ Summary - ETE Robotics", "Please find attached the RFQ summary.", {
            "RFQ_Summary.pdf": pdf_buffer,
            "RFQ_Summary.pptx": ppt_buffer
        })
        st.success("Email sent successfully.")
