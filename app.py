import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

# Charger les données Excel
@st.cache_data

def load_data():
    df_main = pd.read_excel("bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")
    feuilles = {
        "cabine": pd.read_excel("bdd_ht.xlsx", sheet_name="CABINES"),
        "chassis": pd.read_excel("bdd_ht.xlsx", sheet_name="CHASSIS"),
        "caisse": pd.read_excel("bdd_ht.xlsx", sheet_name="CAISSES"),
        "moteur": pd.read_excel("bdd_ht.xlsx", sheet_name="MOTEURS"),
        "frigo": pd.read_excel("bdd_ht.xlsx", sheet_name="FRIGO"),
        "hayon": pd.read_excel("bdd_ht.xlsx", sheet_name="HAYONS")
    }
    return df_main, feuilles

df_main, composants = load_data()

st.title("Générateur de Fiche Technique")

# === Étape 1 : Choix du modèle ===
modeles = df_main["Modele"].dropna().unique()
modele_select = st.selectbox("Choisir un modèle", modeles)

# === Étape 2 : Récupérer toutes les options compatibles avec ce modèle ===
data_filtered = df_main[df_main["Modele"] == modele_select]

def get_unique_options(col):
    return data_filtered[col].dropna().unique()

code_cabine = st.selectbox("Choisir une cabine", get_unique_options("C_Cabine"))
code_chassis = st.selectbox("Choisir un châssis", get_unique_options("C_Chassis"))
code_caisse = st.selectbox("Choisir une caisse", get_unique_options("C_Caisse"))
code_moteur = st.selectbox("Choisir un moteur", get_unique_options("M_moteur"))
code_frigo = st.selectbox("Choisir un groupe frigo", get_unique_options("C_Groupe frigo"))
code_hayon = st.selectbox("Choisir un hayon", get_unique_options("C_Hayon elevateur"))

# === Génération ===
def write_details(ws, df, code, title, start_row):
    if pd.isna(code):
        return start_row
    bloc = df[df[df.columns[0]] == code]
    if bloc.empty:
        return start_row
    ws[f"A{start_row}"] = title
    for i, col in enumerate(bloc.columns):
        ws.cell(row=start_row + 1, column=i + 1, value=col)
        ws.cell(row=start_row + 2, column=i + 1, value=str(bloc[col].values[0]))
    return start_row + 4

if st.button("Générer la fiche technique"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiche Technique"

    ws["A1"] = "Modèle"
    ws["B1"] = modele_select

    row = 3
    row = write_details(ws, composants["cabine"], code_cabine, "Cabine", row)
    row = write_details(ws, composants["chassis"], code_chassis, "Châssis", row)
    row = write_details(ws, composants["caisse"], code_caisse, "Caisse", row)
    row = write_details(ws, composants["moteur"], code_moteur, "Moteur", row)
    row = write_details(ws, composants["frigo"], code_frigo, "Frigo", row)
    row = write_details(ws, composants["hayon"], code_hayon, "Hayon", row)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success("Fiche technique générée avec succès !")
    st.download_button(
        label="📄 Télécharger la fiche technique",
        data=buffer,
        file_name=f"Fiche_{modele_select}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

