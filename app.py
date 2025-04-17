
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# Chargement des données
file = "bdd_ht.xlsx"
df_modeles = pd.read_excel(file, sheet_name="FS_referentiel_produits_std")
df_cabines = pd.read_excel(file, sheet_name="CABINES")
df_chassis = pd.read_excel(file, sheet_name="CHASSIS")
df_caisses = pd.read_excel(file, sheet_name="CAISSES")
df_moteurs = pd.read_excel(file, sheet_name="MOTEURS")
df_frigos = pd.read_excel(file, sheet_name="FRIGO")
df_hayons = pd.read_excel(file, sheet_name="HAYONS")

# Titre et logo
st.markdown("<div style='text-align:center'><img src='https://raw.githubusercontent.com/hibatachfine/Fiche_technique/main/petit_forestier_logo_officiel.png' width='200'></div>", unsafe_allow_html=True)
st.markdown("<h1 style='text-align:center; color:#006400;'>Générateur de Fiche Technique</h1>", unsafe_allow_html=True)
st.markdown("---")

# Sélection modèle
modele = st.selectbox("🛠️ Choisir un modèle", sorted(df_modeles["Modele"].dropna().unique()))
filtres = df_modeles[df_modeles["Modele"] == modele].iloc[0]

# Menus intelligents pour chaque partie
code_cabine = st.selectbox("🚛 Choisir une cabine", df_cabines["C_Cabine"].dropna().unique(), index=df_cabines["C_Cabine"].dropna().tolist().index(filtres["C_Cabine"]) if filtres["C_Cabine"] in df_cabines["C_Cabine"].values else 0)
code_chassis = st.selectbox("🔧 Choisir un châssis", df_chassis["c_chassis"].dropna().unique(), index=df_chassis["c_chassis"].dropna().tolist().index(filtres["c_chassis"]) if filtres["c_chassis"] in df_chassis["c_chassis"].values else 0)
code_caisse = st.selectbox("🚚 Choisir une caisse", df_caisses["c_caisse"].dropna().unique(), index=df_caisses["c_caisse"].dropna().tolist().index(filtres["c_caisse"]) if filtres["c_caisse"] in df_caisses["c_caisse"].values else 0)
code_moteur = st.selectbox("⚙️ Choisir un moteur", df_moteurs["M_moteur"].dropna().unique(), index=df_moteurs["M_moteur"].dropna().tolist().index(filtres["M_moteur"]) if filtres["M_moteur"] in df_moteurs["M_moteur"].values else 0)
code_frigo = st.selectbox("❄️ Choisir un groupe frigo", df_frigos["c_groupe frigo"].dropna().unique(), index=df_frigos["c_groupe frigo"].dropna().tolist().index(filtres["c_groupe frigo"]) if filtres["c_groupe frigo"] in df_frigos["c_groupe frigo"].values else 0)
code_hayon = st.selectbox("🪜 Choisir un hayon", df_hayons["c_hayon elevateur"].dropna().unique(), index=df_hayons["c_hayon elevateur"].dropna().tolist().index(filtres["c_hayon elevateur"]) if filtres["c_hayon elevateur"] in df_hayons["c_hayon elevateur"].values else 0)

# Fonction d’écriture des détails dans Excel
def write_details(df, code, nom_feuille, start_row):
    bloc = df[df.iloc[:, 0] == code]
    wb = load_workbook("bdd_ht.xlsx")
    ws = wb[nom_feuille]
    for col_idx, col in enumerate(bloc.columns):
        ws.cell(row=start_row, column=col_idx + 1, value=str(bloc[col].values[0]))
    return wb

# Génération fiche technique
if st.button("📄 Générer la fiche technique"):
    wb = load_workbook("bdd_ht.xlsx")

    write_details(df_cabines, code_cabine, "CABINES", 10)
    write_details(df_chassis, code_chassis, "CHASSIS", 10)
    write_details(df_caisses, code_caisse, "CAISSES", 10)
    write_details(df_moteurs, code_moteur, "MOTEURS", 10)
    write_details(df_frigos, code_frigo, "FRIGO", 10)
    write_details(df_hayons, code_hayon, "HAYONS", 10)

    # Génération du fichier
    output = BytesIO()
    wb.save(output)
    st.success("✅ Fiche technique générée avec succès !")
    st.download_button("📥 Télécharger la fiche Excel", data=output.getvalue(), file_name=f"Fiche_Technique_{modele}.xlsx")
