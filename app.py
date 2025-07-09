import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

# Authentification simple
def check_password():
    def password_entered():
        if st.session_state["password"] == "FT.petitforestier":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Mot de passe", type="password", on_change=password_entered, key="password")
        st.stop()
    elif not st.session_state["password_correct"]:
        st.text_input("Mot de passe", type="password", on_change=password_entered, key="password")
        st.error("Mot de passe incorrect")
        st.stop()

check_password()

# Titre
st.image("petit_forestier_logo_officiel.png", width=700)
st.markdown("<h1 style='color:#057A20;'>Générateur de Fiches Techniques</h1>", unsafe_allow_html=True)
st.markdown("---")

# Chargement des données
try:
    df = pd.read_excel("bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")
    cabine_df = pd.read_excel("bdd_ht.xlsx", sheet_name="CABINES")
    chassis_df = pd.read_excel("bdd_ht.xlsx", sheet_name="CHASSIS")
    caisse_df = pd.read_excel("bdd_ht.xlsx", sheet_name="CAISSES")
    moteur_df = pd.read_excel("bdd_ht.xlsx", sheet_name="MOTEURS")
    frigo_df = pd.read_excel("bdd_ht.xlsx", sheet_name="FRIGO")
    hayon_df = pd.read_excel("bdd_ht.xlsx", sheet_name="HAYONS")
except Exception as e:
    st.error(f"Erreur de chargement : {e}")
    st.stop()

# Sélections utilisateur
code_pays = st.selectbox("Code pays", sorted(df["Code_Pays"].dropna().unique()))
df_filtered = df[df["Code_Pays"] == code_pays]

marque = st.selectbox("Marque", sorted(df_filtered["Marque"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Marque"] == marque]

modele = st.selectbox("Modèle", sorted(df_filtered["Modele"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Modele"] == modele]

code_pf = st.selectbox("Code PF", sorted(df_filtered["Code_PF"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Code_PF"] == code_pf]

code_cabine = st.selectbox("Cabine", df_filtered["C_Cabine"].dropna().unique())
code_chassis = st.selectbox("Châssis", df_filtered["C_Chassis"].dropna().unique())
code_caisse = st.selectbox("Caisse", df_filtered["C_Caisse"].dropna().unique())
code_moteur = st.selectbox("Moteur", df_filtered["M_Moteur"].dropna().unique())
code_frigo = st.selectbox("Groupe Frigorifique", df_filtered["C_Groupe Frigorifique"].dropna().unique())
code_hayon = st.selectbox("Hayon", df_filtered["C_Hayon"].dropna().unique())

# Fonctions
def get_criteria_list(df, code, code_column):
    row = df[df[code_column] == code]
    if row.empty:
        return []
    row = row.iloc[0].dropna()
    exclude = [code_column, 'Produit (P) / Option (O)']
    return [str(val).strip() for col, val in row.items() if col not in exclude and str(val).strip().lower() != 'nan' and str(val).strip() != '']

def insert_criteria(ws, start_cell, criteria_list):
    col_letter = ''.join(filter(str.isalpha, start_cell))
    start_row = int(''.join(filter(str.isdigit, start_cell)))
    for i, crit in enumerate(criteria_list):
        ws[f"{col_letter}{start_row + i}"] = crit

# Génération du fichier
def generate_filled_ft():
    wb = load_workbook("Modèle FT.xlsx")
    ws = wb.active

    # Infos générales
    ws["E8"] = code_pays
    ws["E9"] = marque
    ws["E10"] = modele
    ws["E11"] = code_pf

    # Critères sous composants
    insert_criteria(ws, "E50", get_criteria_list(cabine_df, code_cabine, "C_Cabine"))
    insert_criteria(ws, "F50", get_criteria_list(moteur_df, code_moteur, "M_moteur"))
    insert_criteria(ws, "G50", get_criteria_list(chassis_df, code_chassis, "C_Chassis"))
    insert_criteria(ws, "H50", get_criteria_list(caisse_df, code_caisse, "C_Caisse"))
    insert_criteria(ws, "I50", get_criteria_list(frigo_df, code_frigo, "C_Groupe Frigorifique"))
    insert_criteria(ws, "J50", get_criteria_list(hayon_df, code_hayon, "C_Hayon"))

    # Export final
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Téléchargement
st.download_button(
    label="📥 Télécharge
