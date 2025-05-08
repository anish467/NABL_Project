import streamlit as st
import openpyxl
import random
from io import BytesIO
import openpyxl.utils.cell as utils

# Define weight options in grams
WEIGHT_OPTIONS = {
    "1mg": 0.001, "2mg": 0.002, "5mg": 0.005, "10mg": 0.01, "20mg": 0.02, "50mg": 0.05,
    "100mg": 0.1, "200mg": 0.2, "500mg": 0.5,
    "1g": 1.0, "2g": 2.0, "5g": 5.0, "10g": 10.0, "20g": 20.0, "50g": 50.0,
    "100g": 100.0, "200g": 200.0, "500g": 500.0,
    "1kg": 1000.0, "2kg": 2000.0, "5kg": 5000.0, "10kg": 10000.0, "20kg": 20000.0, "50kg": 50000.0
}

weights_dict = {
    "1mg": 0.0010013,
    "2mg": 0.0020009,
    "2mgs": 0.0020005,
    "5mg": 0.0050013,
    "10mg": 0.0100004,
    "20mg": 0.0200011,
    "20mgs": 0.0200013,
    "50mg": 0.0500016,
    "100mg": 0.1000017,
    "200mg": 0.2000024,
    "200mgs": 0.2000028,
    "500mg": 0.5000039,
    "1g": 0.9999987,
    "2g": 2.0000020,
    "2gs": 2.0000030,
    "5g": 5.000002,
    "10g": 10.000011,
    "20g": 20.000008,
    "20gs": 20.000013,
    "50g": 50.000014,
    "100g": 100.000021,
    "200g": 200.00005,
    "200gs": 200.00006,
    "500g": 500.00016,
    "1kg": 1000.0002,
    "2kg": 2000.0005,
    "5kg": 5000.0012,
    "10kg": 10000.0073,
    "20kg": 19999.996,
    "50kg": 50000.015
}

pos1 = [
    (0, 0, 1, 0),
    (0, 1, 0, 0),
    (0, 1, 1, 1),
    (1, 1, 1, 0),
    (1, 1, 2, 1),
    (1, 2, 1, 1),
    (1, 2, 2, 2),
    (2, 2, 2, 1)
]

pos0 = [
    (0, 0, 1, 1),
    (0, 1, 0, 1),
    (1, 0, 1, 0),
    (1, 1, 0, 0),
    (1, 1, 1, 1),
    (1, 1, 2, 2),
    (1, 2, 1, 2),
    (2, 1, 2, 1),
    (2, 2, 1, 1),
    (2, 2, 2, 2)
]

neg1 = [
    (-2, -2, -1, -2),
    (-2, -1, -2, -2),
    (-2, -1, -1, -1),
    (-1, -2, 0, -2),
    (-1, -1, -1, -2),
    (-1, -1, 0, -1),
    (-1, 0, -1, -1),
    (-1, 0, 0, 0),
    ( 0, 0, 0, -1)
]

neg0 = [
    (-2, -2, -2, -2),
    (-2, -2, -1, -1),
    (-2, -1, -2, -1),
    (-1, -2, -1, -2),
    (-1, -1, -2, -2),
    (-1, -1, -1, -1),
    (-1, -1, 0, 0),
    (-1, -1, 1, 1),
    (-1, 0, -1, 0),
    (-1, 0, 0, 1),
    ( 0, -1, 0, -1),
    ( 0, 0, -1, -1),
    ( 0, 0, 0, 0),
    ( 1, 0, 0, -1)
]

negpos1 = [
    (0, -1, 1, 1),
    (0, 0, 0, 1),
    (1, 0, 0, 0),
    (1, 0, 1, 1),
    (1, 1, 0, 1),
    (1, 1, 1, 2),
    (2, 1, 0, 0),
    (2, 1, 1, 1),
    (2, 1, 2, 2),
    (2, 2, 1, 2)
]

pos2 = [
    (-1, 1, 1, 1),
    ( 0, 1, 1, 0),
    ( 0, 1, 2, 1),
    ( 0, 2, 1, 1),
    ( 1, 1, 2, 0),
    ( 1, 2, 1, 0),
    ( 1, 2, 2, 1)
]

negneg1 = [
    (-2, -2, -2, -1),
    (-1, -2, -2, -2),
    (-1, -2, -1, -1),
    (-1, -1, -2, -1),
    (-1, -1, -1, 0),
    ( 0, -1, -2, -2),
    ( 0, -1, -1, -1),
    ( 0, -1, 0, 0),
    ( 0, 0, -1, 0)
]

neg2 = [
    (-2, -1, -1, -2),
    (-2, -1, 0, -1),
    (-2, 0, -1, -1),
    (-1, -1, 0, -2),
    (-1, -1, 1, -1),
    (-1, 0, -1, -2),
    (-1, 0, 0, -1),
    (-1, 0, 1, 0),
    (-1, 1, 0, 0),
    ( 0, 0, 1, -1),
    ( 0, 1, 0, -1)
]

negpos2= [
    ( 1, 0, 0, 1),
    ( 1, 0, 1, 2),
    ( 1, 1, 0, 2),
    ( 2, 0, 1, 1),
    ( 2, 1, 0, 1),
    ( 2, 1, 1, 2)
]

negneg2 = [
    (-1, -2, -2, -1),
    (-1, -2, -1, 0),
    (-1, -1, -2, 0),
    (-1, -1, -1, 1),
    ( 0, -2, -1, -1),
    ( 0, -1, -2, -1),
    ( 0, -1, -1, 0),
    ( 0, -1, 0, 1),
    ( 0, 0, -1, 1),
    ( 1, -1, -1, -1),
    ( 1, -1, 0, 0),
    ( 1, 0, -1, 0)
]

combination_list = [
    ("pos1", "pos0"),
    ("pos1", "negpos1"),
    ("pos0", "pos0"),
    ("pos0", "negpos1"),
    ("pos0", "pos1"),
    ("neg1", "neg0"),
    ("neg1", "negneg1"),
    ("neg0", "negneg1"),
    ("neg0", "neg0"),
    ("neg0", "neg1"),
    ("negpos1", "pos1"),
    ("negneg1", "neg1"),
]




# ABBA generation function as provided
def generate_abba_readings(target_value, reference_weight):
    random_pair = random.choice(combination_list)
    tup_dict = {
    'neg0': neg0,
    'neg1': neg1,
    'neg2': neg2,
    'negneg1': negneg1,
    'negneg2': negneg2,
    'negpos1': negpos1,
    'negpos2': negpos2,
    'pos0': pos0,
    'pos1': pos1,
    'pos2': pos2
}
    cycle1 = random.choice(tup_dict[random_pair[0]])
    cycle2 = random.choice(tup_dict[random_pair[1]])
    decimal_part = str(target_value).split('.')[-1]
    decimal_places = len(decimal_part)
    x = 10 ** -decimal_places
    return [
        (x*float(cycle1[0]) + reference_weight),
        (x*float(cycle1[1]) + target_value),
        (x*float(cycle1[2]) + target_value),
        (x*float(cycle1[3]) + reference_weight),
        (x*float(cycle2[0]) + reference_weight),
        (x*float(cycle2[1]) + target_value),
        (x*float(cycle2[2]) + target_value),
        (x*float(cycle2[3]) + reference_weight),
    ]


# Insert ABBA readings into the Excel sheet
def insert_abba_values(sheet, readings, start_row, decimal_places):
    for i in range(2):
        for j in range(4):
            col_letter = chr(ord('H') + j)
            cell = f"{col_letter}{start_row + i}"
            sheet[cell] = round(readings[i * 4 + j], decimal_places)

# Streamlit app interface
st.title("⚖️ ABBA Cycle Weight Calibration Tool")

uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])
num_weights = st.number_input("🔢 How many weights to calibrate?", min_value=1, max_value=50, step=1)

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file)
    sheet_names = "INTERMEDIATE"

    weight_data = []
    reference_weights = []

    st.subheader("📋 Input for Each Weight")

    for i in range(num_weights):
        with st.expander(f"➕ Weight #{i+1}"):
            col1, col2, col3 , col4 = st.columns(4)
            with col2:
                identi = st.text_input("Identification Mark",value="",key=f"identi_{i}")
            with col1:
                denomination = st.selectbox(f"Denomination", list(WEIGHT_OPTIONS.keys()), key=f"denom_{i}")
            with col3:
                target_weight = st.number_input(f"Target Weight (grams)", key=f"target_{i}", format="%.6f")
            with col4:
                start_cell = st.text_input(f"Start Cell", key=f"cell_{i}", value=f"H{34 + (i * 2)}")

            if denomination and target_weight and start_cell:
                reference_weight = weights_dict[denomination]
                print(reference_weight)
                decimal_places = len(str(target_weight).split('.')[-1])
                reference_weights.append(reference_weight)
                weight_data.append({
                    "denomination": denomination,
                    "identi": identi,
                    "target_weight": target_weight,
                    "start_cell": start_cell,
                    "reference_weight": reference_weight,
                    "decimal_places": decimal_places
                })

    if weight_data and st.button("🚀 Insert ABBA Values into Excel"):
        sheet = wb[sheet_names]

        for i, wd in enumerate(weight_data):
            row_index = int(utils.coordinate_from_string(wd["start_cell"])[1])
            sheet[f"D{row_index}"] = wd["denomination"]
            sheet[f"C{row_index}"] = wd["identi"]

            ref_weight = reference_weights[i]
            abba_values = generate_abba_readings(wd["target_weight"], ref_weight)
            insert_abba_values(sheet, abba_values, row_index, wd["decimal_places"])

            st.success(f"ABBA values for {wd['denomination']} inserted at {wd['start_cell']}")

        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Download Modified Excel File",
            data=output.getvalue(),
            file_name="modified_abba_weights.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
