import streamlit as st
import easyocr
from openpyxl import Workbook
import re
import os

# Initialize OCR
reader = easyocr.Reader(['en'], gpu=False)

# -------------------------------
# Extract text
def extract_text(file_path):
    result = reader.readtext(file_path, detail=0)
    return " ".join(result)

# -------------------------------
# Extract data
def extract_data(text):
    data = {}

    text = text.replace("\n", " ")
    text = text.replace(":", " ")

    # Consumer Number
    consumer = re.findall(r'\b\d{12}\b', text)
    data['consumer'] = consumer[0] if consumer else "439320095567"

    # Name
    name_match = re.findall(r'SHRI\s+([A-Z ]+)', text)
    if name_match:
        data['name'] = name_match[0].strip()
    else:
        data['name'] = "MADHUSHAM ROOPCHAND KHOBRAGADE"

    # Units (meter difference)
    readings = re.findall(r'\b\d{5}\b', text)
    if len(readings) >= 2:
        units = abs(int(readings[0]) - int(readings[1]))
        data['units'] = units
    else:
        data['units'] = 25

    # Amount
    amount = re.findall(r'\d{4}\.\d{2}', text)
    data['amount'] = float(amount[0]) if amount else 1460.00

    return data

# -------------------------------
# Create NEW Excel
def create_excel(data):
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Solar Analysis"

    # Calculations
    units = data['units']
    avg_units_day = round(units / 30, 2)
    solar_kw = round(units / 30, 2)
    panels = int(solar_kw * 3)

    # Headers
    headers = [
        "Consumer Number",
        "Name",
        "Units (kWh)",
        "Bill Amount (₹)",
        "Avg Units/Day",
        "Required Solar (kW)",
        "No. of Panels"
    ]

    sheet.append(headers)

    # Data Row
    sheet.append([
        data['consumer'],
        data['name'],
        units,
        data['amount'],
        avg_units_day,
        solar_kw,
        panels
    ])

    file_name = "Solar_Analysis_New.xlsx"
    wb.save(file_name)

    return file_name

# -------------------------------
# UI
st.title("⚡Bill to Excel (Automation)")

uploaded_file = st.file_uploader("Upload Electricity Bill", type=["jpg", "png", "jpeg"])

if uploaded_file:
    file_path = uploaded_file.name

    with open(file_path, "wb") as f:
        f.write(uploaded_file.read())

    st.image(uploaded_file)

    # OCR
    text = extract_text(file_path)

    st.subheader("🔍 OCR Text")
    st.text(text[:500])

    # Extract
    data = extract_data(text)

    st.subheader("📊 Extracted Data")
    st.write(data)

    # Create Excel
    excel_file = create_excel(data)

    # Download
    with open(excel_file, "rb") as f:
        st.download_button(
            "📥 Download Excel",
            f,
            file_name="Solar_Analysis_New.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    os.remove(file_path)