import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import find_peaks
from io import BytesIO
from openpyxl import Workbook

# Fonction pour détecter les pics locaux
def detect_peaks(wavenumber, absorbance, threshold=0.02):
    peaks, _ = find_peaks(absorbance, height=threshold)
    peak_positions = wavenumber[peaks]
    peak_heights = absorbance[peaks]
    return peak_positions, peak_heights

# Charger la table de correspondance Additifs - Pics
def load_reference_table():
    reference = {
        "T-622": [1736],
        "C-944": [1568, 1530],
        "C-3346": [1540],
        "Chimassorb 119": [1530],
        "I-1010": [1745],
        "I-1076": [1740],
        "I-168": [1212, 1194, 850],
        "W-399": [1205],
        "Cyanox": [1790, 1705],
        "Slip (KE, KS, KU, KB)": [1640],
        "Slip (EBS)": [1640, 1554],
        "lauryl amide (antistat)": [1620],
        "Ca(St)2": [1578, 1540],
        "Zn(St)2": [1539],
        "Talc": [1020, 670, 465, 451],
        "Superfloss": [1087, 800, 475],
        "CaCO3": [875, 1450, 1800],
        "CaSO4": [1155, 1125, 675, 613, 595],
        "EVA": [1740, 1238, 1020],
        "Degraded carbonyl": [1718],
        "Butene comonomer": [1378, 770],
        "Hexene comonomer": [1377, 895],
        "Octene comonomer": [1377],
        "PE": [2940, 1460, 730, 720],
        "PP": [1167, 997, 972, 841],
        "PIB": [1230, 950, 923],
        "Polybutylene": [1380, 760],
        "Polyamide": [1636, 1540, 3300],
        "PET": [1730, 1248, 1110],
        "EVOH": [1100, 850, 3350],
        "Surlyn (Na)": [1699],
        "Surlyn (Zn)": [1699, 1585],
        "Acrylics": [1735, 1165],
        "PS": [756, 700, 1600],
        "Silicone": [1262, 1020, 799]
    }
    return reference

# Fonction de matching

def match_substances(selected_peaks, reference_table, tolerance):
    results = []
    for substance, ref_peaks in reference_table.items():
        for ref_peak in ref_peaks:
            for user_peak in selected_peaks:
                if abs(user_peak - ref_peak) <= tolerance:
                    results.append({
                        "Substance": substance,
                        "Peak Detected (cm-1)": user_peak,
                        "Reference Peak (cm-1)": ref_peak,
                        "Justification": f"{substance} detected via peak at {user_peak} cm-1 (ref: {ref_peak} cm-1)"
                    })
    return pd.DataFrame(results)

# Fonction pour exporter en Excel
def export_to_excel(peaks_df, match_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        match_df.to_excel(writer, index=False, sheet_name='Detection Results')
        peaks_df.to_excel(writer, index=False, sheet_name='Raw Peaks')
    output.seek(0)
    return output

# Streamlit UI
st.title("FTIR Peak Analysis - Additives Detection")

uploaded_file = st.file_uploader("Upload your FTIR CSV file", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.write("### Raw Data Preview:")
    st.dataframe(df.head())

    # Supposer que les 2 colonnes principales sont les 2 premières
    wavenumber = df.iloc[:,0].values
    absorbance = df.iloc[:,1].values

    threshold = st.slider("Peak Detection Threshold (Absorbance)", 0.01, 0.1, 0.02, 0.01)
    tolerance = st.slider("Tolerance (cm-1)", 1, 10, 5, 1)

    peak_positions, peak_heights = detect_peaks(wavenumber, absorbance, threshold)

    # Afficher graphique
    fig, ax = plt.subplots()
    ax.plot(wavenumber, absorbance, color='red')
    ax.scatter(peak_positions, peak_heights, color='blue')
    ax.set_xlabel("Wavenumber (cm-1)")
    ax.set_ylabel("Absorbance")
    ax.set_title("FTIR Spectrum with Detected Peaks")
    ax.invert_xaxis()
    st.pyplot(fig)

    peaks_df = pd.DataFrame({"Wavenumber (cm-1)": peak_positions, "Absorbance": peak_heights})

    st.write("### Select Peaks to Analyze:")
    selected_peaks = st.multiselect("Select Peaks (cm-1)", peak_positions.tolist(), default=peak_positions.tolist())

    if selected_peaks:
        reference_table = load_reference_table()
        match_df = match_substances(selected_peaks, reference_table, tolerance)

        st.write("### Detection Results:")
        st.dataframe(match_df)

        excel_data = export_to_excel(peaks_df, match_df)
        st.download_button(
            label="Download Results as Excel",
            data=excel_data,
            file_name="ftir_detection_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Please select at least one peak to analyze.")
