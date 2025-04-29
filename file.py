import streamlit as st
import pandas as pd
import tempfile
from src import ps_e

st.set_page_config(layout="wide")

rainbow_title = """
<h1 style='text-align: center; background: linear-gradient(to right, red, orange, yellow, green, blue, indigo, violet);
            -webkit-background-clip: text; color: transparent; font-size: 40px;'>
    MEP Calculator
</h1>
"""
st.markdown(rainbow_title, unsafe_allow_html=True)

st.markdown("""
<div style='text-align: left; font-size:18px; background-color: #f0f2f6; padding: 15px; border-radius: 8px;'>
<b>About the MEP Calculator</b><br><br>
The <b>MEP Calculator</b> is a tool to help working on energy-efficient building projects, such as LEED-certified projects. By uploading four SIM files (representing 0°, 90°, 180°, and 270° rotations).<br><br>

<b>Key Features:</b><br>
- Upload and process four SIM files representing different rotations.<br>
- Automatically update MEP values in an Excel sheet based on the performance data extracted from the SIM files.<br>
- Display the updated table and allow downloading of the modified Excel file.<br>
- Useful for energy-efficient building design, particularly for LEED projects.<br>
</div>
""", unsafe_allow_html=True)

sim_files = st.file_uploader("Choose 4 SIM files", type=["sim"], accept_multiple_files=True)
fixed_excel_path = r"database/v4_Minimum_Energy_Performance_Calculator-v06.xlsm"

if fixed_excel_path:
    df_full = pd.read_excel(fixed_excel_path, sheet_name="Performance_Outputs_1", header=None)
    table_rows = df_full[df_full.apply(lambda row: row.astype(str).str.contains('Table:', na=False).any(), axis=1)]
    table_names = table_rows.apply(lambda row: row[row.astype(str).str.contains('Table:')].values[0].split("Table:")[1].strip(), axis=1).tolist()

    st.subheader("Select Tables to Update")
    selected_tables = []
    table_name = None
    for table in table_names:
        if st.checkbox(f"{table}"):
            words = table.split()[:3]
            capitalized_words = [word.capitalize() for word in words]
            table_ = ''.join(capitalized_words) + '.xlsx'
            st.write(table_)
            table_name = r'tables/' + table_
            selected_tables.append(table)

if st.button("Process Files"):
    if len(sim_files) != 4:
        st.warning("Please upload exactly 4 SIM files.")
    elif table_name is None:
        st.warning("Please select at least one table to update.")
    else:
        df = pd.read_excel(table_name) 
        st.write(df)

        end_use_map = {
            "Interior lighting": "LIGHTS",
            "Exterior lighting": "EXT USAGE",
            "Space heating": "SPACE_HEATING",
            "Space cooling": "SPACE_COOLING",
            "Pumps": "PUMPS & AUX",
            "Heat rejection": "HEAT_REJECT",
            "Fans - interior ventilation": "VENT FANS",
            "Service water heating": "DOMEST HOT WTR"
        }

        pse_dfs = []
        rotation_labels = [
            'Baseline 0° rotation',
            'Baseline 90° rotation',
            'Baseline 180° rotation',
            'Baseline 270° rotation'
        ]

        for i, sim_file in enumerate(sim_files):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".sim") as temp_file:
                temp_file.write(sim_file.read())
                temp_file_path = temp_file.name
                
            pse_df = ps_e.get_PSE_report(temp_file_path)

            pse_df_int_light_kwh = pse_df['LIGHTS'][0]
            pse_df_int_light_kw = pse_df['LIGHTS'][1]
            pse_df_ext_light_kwh = pse_df['EXT USAGE'][0]
            pse_df_ext_light_kw = pse_df['EXT USAGE'][1]
            pse_df_cool_kwh = pse_df['SPACE_COOLING'][0]
            pse_df_cool_kw = pse_df['SPACE_COOLING'][1]
            pse_df_pumps_kwh = pse_df['PUMPS & AUX'][0]
            pse_df_pumps_kw = pse_df['PUMPS & AUX'][1]
            pse_df_heat_reject_kwh = pse_df['HEAT_REJECT'][0]
            pse_df_heat_reject_kw = pse_df['HEAT_REJECT'][1]
            pse_df_fans_kwh = pse_df['VENT FANS'][0]
            pse_df_fans_kw = pse_df['VENT FANS'][1]
            pse_df_wtr_kwh = pse_df['DOMEST HOT WTR'][0]
            pse_df_wtr_kw = pse_df['DOMEST HOT WTR'][1]

            col = rotation_labels[i]
            df[col][0] = float(pse_df_int_light_kwh)
            df[col][1] = float(pse_df_int_light_kw)
            df[col][2] = float(pse_df_ext_light_kwh)
            df[col][3] = float(pse_df_ext_light_kw)
            df[col][6] = float(pse_df_cool_kwh)
            df[col][7] = float(pse_df_cool_kw)
            df[col][8] = float(pse_df_pumps_kwh)
            df[col][9] = float(pse_df_pumps_kw)
            df[col][10] = float(pse_df_heat_reject_kwh)
            df[col][11] = float(pse_df_heat_reject_kw)
            df[col][12] = float(pse_df_fans_kwh)
            df[col][13] = float(pse_df_fans_kw)
            df[col][16] = float(pse_df_wtr_kwh)
            df[col][17] = float(pse_df_wtr_kw)

        cols = [
            'Baseline 0° rotation',
            'Baseline 90° rotation',
            'Baseline 180° rotation',
            'Baseline 270° rotation'
        ]
        df = df.iloc[:, :-4]
        df = df.iloc[:, :-1]
        df = df.drop(df.columns[1], axis=1)
        df[cols] = df[cols].apply(pd.to_numeric, errors='coerce')
        df['Baseline Design Total (Average of 4 rotations)'] = df[cols].sum(axis=1) / 4

        st.success("Files processed and CSV updated successfully!")
        st.dataframe(df)

        st.download_button(
            label="Download Modified CSV",
            data=df.to_csv(index=False),
            file_name="modified_output.csv",
            mime="text/csv"
        )