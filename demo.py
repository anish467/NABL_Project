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

pos1 = [
    (-1, 0, 1, 1),
    (-1, 1, 0, 1),
    (0, 0, 1, 0),
    (0, 0, 2, 1),
    (0, 1, 0, 0),
    (0, 1, 1, 1),
    (0, 1, 2, 2),
    (0, 2, 0, 1),
    (0, 2, 1, 2),
    (1, 0, 1, -1),
    (1, 0, 2, 0),
    (1, 1, 0, -1),
    (1, 1, 1, 0),
    (1, 1, 2, 1),
    (1, 2, 0, 0),
    (1, 2, 1, 1),
    (1, 2, 2, 2),
    (2, 1, 2, 0),
    (2, 2, 1, 0),
    (2, 2, 2, 1)
]

pos0 = [
    (0, 0, 1, 1),
    (0, 0, 2, 2),
    (0, 1, 0, 1),
    (0, 1, 1, 2),
    (0, 2, 0, 2),
    (1, 0, 1, 0),
    (1, 0, 2, 1),
    (1, 1, 0, 0),
    (1, 1, 1, 1),
    (1, 1, 2, 2),
    (1, 2, 0, 1),
    (1, 2, 1, 2),
    (2, 0, 2, 0),
    (2, 1, 1, 0),
    (2, 1, 2, 1),
    (2, 2, 0, 0),
    (2, 2, 1, 1),
    (2, 2, 2, 2)
]

neg1 = [
    (-2, -2, -1, -2),
    (-2, -2,  0, -1),
    (-2, -1, -2, -2),
    (-2, -1, -1, -1),
    (-2, -1,  0,  0),
    (-2,  0, -2, -1),
    (-2,  0, -1,  0),
    (-1, -2,  0, -2),
    (-1, -1, -1, -2),
    (-1, -1,  0, -1),
    (-1, -1,  1,  0),
    (-1,  0, -2, -2),
    (-1,  0, -1, -1),
    (-1,  0,  0,  0),
    (-1,  1, -1,  0),
    ( 0, -1,  0, -2),
    ( 0, -1,  1, -1),
    ( 0,  0, -1, -2),
    ( 0,  0,  0, -1),
    ( 0,  1, -1, -1)
]

neg0 = [
    (-2, -2, -2, -2),
    (-2, -2, -1, -1),
    (-2, -2,  0,  0),
    (-2, -1, -2, -1),
    (-2, -1, -1,  0),
    (-2,  0, -2,  0),
    (-1, -2, -1, -2),
    (-1, -2,  0, -1),
    (-1, -1, -2, -2),
    (-1, -1, -1, -1),
    (-1, -1,  0,  0),
    (-1, -1,  1,  1),
    (-1,  0, -2, -1),
    (-1,  0, -1,  0),
    (-1,  0,  0,  1),
    (-1,  1, -1,  1),
    ( 0, -2,  0, -2),
    ( 0, -1, -1, -2),
    ( 0, -1,  0, -1),
    ( 0, -1,  1,  0),
    ( 0,  0, -2, -2),
    ( 0,  0, -1, -1),
    ( 0,  0,  0,  0),
    ( 0,  1, -1,  0),
    ( 1, -1,  1, -1),
    ( 1,  0,  0, -1),
    ( 1,  1, -1, -1)
]

negpos1 = [
    (0, -1, 1, 1),
    (0,  0, 0, 1),
    (0,  0, 1, 2),
    (0,  1, -1, 1),
    (0,  1, 0, 2),
    (1, -1, 1, 0),
    (1,  0, 0, 0),
    (1,  0, 1, 1),
    (1,  0, 2, 2),
    (1,  1, -1, 0),
    (1,  1, 0, 1),
    (1,  1, 1, 2),
    (1,  2, 0, 2),
    (2,  0, 1, 0),
    (2,  0, 2, 1),
    (2,  1, 0, 0),
    (2,  1, 1, 1),
    (2,  1, 2, 2),
    (2,  2, 0, 1),
    (2,  2, 1, 2)
]

pos2 = [
    (-1, 0, 2, 1),
    (-1, 1, 1, 1),
    (-1, 2, 0, 1),
    ( 0, 0, 2, 0),
    ( 0, 1, 1, 0),
    ( 0, 1, 2, 1),
    ( 0, 2, 0, 0),
    ( 0, 2, 1, 1),
    ( 0, 2, 2, 2),
    ( 1, 0, 2, -1),
    ( 1, 1, 1, -1),
    ( 1, 1, 2,  0),
    ( 1, 2, 0, -1),
    ( 1, 2, 1,  0),
    ( 1, 2, 2,  1),
    ( 2, 2, 2,  0)
]

negneg1 = [
    (-2, -2, -2, -1),
    (-2, -2, -1,  0),
    (-2, -1, -2,  0),
    (-1, -2, -2, -2),
    (-1, -2, -1, -1),
    (-1, -2,  0,  0),
    (-1, -1, -2, -1),
    (-1, -1, -1,  0),
    (-1, -1,  0,  1),
    (-1,  0, -2,  0),
    (-1,  0, -1,  1),
    ( 0, -2, -1, -2),
    ( 0, -2,  0, -1),
    ( 0, -1, -2, -2),
    ( 0, -1, -1, -1),
    ( 0, -1,  0,  0),
    ( 0,  0, -2, -1),
    ( 0,  0, -1,  0),
    ( 1, -1,  0, -1),
    ( 1,  0, -1, -1)
]

neg2 = [
    (-2, -2,  0, -2),
    (-2, -1, -1, -2),
    (-2, -1,  0, -1),
    (-2, -1,  1,  0),
    (-2,  0, -2, -2),
    (-2,  0, -1, -1),
    (-2,  0,  0,  0),
    (-2,  1, -1,  0),
    (-1, -1,  0, -2),
    (-1, -1,  1, -1),
    (-1,  0, -1, -2),
    (-1,  0,  0, -1),
    (-1,  0,  1,  0),
    (-1,  1, -1, -1),
    (-1,  1,  0,  0),
    ( 0, -1,  1, -2),
    ( 0,  0,  0, -2),
    ( 0,  0,  1, -1),
    ( 0,  1, -1, -2),
    ( 0,  1,  0, -1)
]

negpos2= [
    ( 0, -1,  1,  2),
    ( 0,  0,  0,  2),
    ( 0,  1, -1,  2),
    ( 1, -1,  1,  1),
    ( 1,  0,  0,  1),
    ( 1,  0,  1,  2),
    ( 1,  1, -1,  1),
    ( 1,  1,  0,  2),
    ( 2, -1,  1,  0),
    ( 2,  0,  0,  0),
    ( 2,  0,  1,  1),
    ( 2,  0,  2,  2),
    ( 2,  1, -1,  0),
    ( 2,  1,  0,  1),
    ( 2,  1,  1,  2),
    ( 2,  2,  0,  2)
]

negneg2 = [
    (-2, -2, -2,  0),
    (-1, -2, -2, -1),
    (-1, -2, -1,  0),
    (-1, -2,  0,  1),
    (-1, -1, -2,  0),
    (-1, -1, -1,  1),
    (-1,  0, -2,  1),
    ( 0, -2, -2, -2),
    ( 0, -2, -1, -1),
    ( 0, -2,  0,  0),
    ( 0, -1, -2, -1),
    ( 0, -1, -1,  0),
    ( 0, -1,  0,  1),
    ( 0,  0, -2,  0),
    ( 0,  0, -1,  1),
    ( 1, -2,  0, -1),
    ( 1, -1, -1, -1),
    ( 1, -1,  0,  0),
    ( 1,  0, -2, -1),
    ( 1,  0, -1,  0)
]

combination_list = [
    ("pos1", "pos0"),
    ("pos1", "pos1"),
    ("pos1", "pos2"),
    ("pos0", "pos0"),
    ("pos0", "negpos1"),
    ("pos0", "pos1"),
    ("neg1", "neg0"),
    ("neg1", "neg1"),
    ("neg1", "neg2"),
    ("neg0", "negneg1"),
    ("neg0", "neg0"),
    ("neg0", "neg1"),
    ("negpos1", "negpos1"),
    ("negpos1", "pos0"),
    ("negpos1", "negpos2"),
    ("pos2", "negpos1"),
    ("pos2", "pos2"),
    ("negneg1", "neg2"),
    ("negneg1", "negneg1"),
    ("negneg1", "neg0"),
    ("neg2", "neg2"),
    ("neg2", "neg1"),
    ("negpos2", "negpos2"),
    ("negpos2", "negpos1"),
    ("negneg2", "negneg1"),
    ("negneg2", "negneg1")
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
            col1, col2, col3 = st.columns(3)
            with col1:
                denomination = st.selectbox(f"Denomination #{i+1}", list(WEIGHT_OPTIONS.keys()), key=f"denom_{i}")
            with col2:
                target_weight = st.number_input(f"Target Weight #{i+1} (grams)", key=f"target_{i}", format="%.6f")
            with col3:
                start_cell = st.text_input(f"Start Cell #{i+1} (e.g., B2)", key=f"cell_{i}", value=f"H{34 + (i * 2)}")

            if denomination and target_weight and start_cell:
                reference_weight = WEIGHT_OPTIONS[denomination]
                decimal_places = len(str(target_weight).split('.')[-1])
                reference_weights.append(reference_weight)
                weight_data.append({
                    "denomination": denomination,
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
