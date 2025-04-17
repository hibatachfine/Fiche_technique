import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image

# Titre et logo
st.image("petit_forestier_logo_officiel.png", width=700)
st.markdown("<h1 style='color:#057A20;'>Générateur de Fiches Techniques</h1>", unsafe_allow_html=True)
st.markdown("---")

# Chargement des données
try:
    df = pd.read_excel("bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")
except Exception as e:
    st.error(f"Erreur lors du chargement du fichier Excel : {e}")
    st.stop()

# Colonnes attendues
required_columns = ["Modele", "C_Cabine", "C_Chassis", "C_Caisse", "M_moteur", "C_Groupe frigo", "C_Hayon elevateur"]
if not all(col in df.columns for col in required_columns):
    st.error("Colonnes manquantes dans le fichier Excel: " + ", ".join(required_columns))
    st.stop()

# Sélections
modele = st.selectbox("Choisir un modèle", sorted(df["Modele"].dropna().unique()))
df_filtered = df[df["Modele"] == modele]

code_cabine = st.selectbox("Choisir une cabine", df_filtered["C_Cabine"].dropna().unique())
code_chassis = st.selectbox("Choisir un châssis", df_filtered["C_Chassis"].dropna().unique())
code_caisse = st.selectbox("Choisir une caisse", df_filtered["C_Caisse"].dropna().unique())
code_moteur = st.selectbox("Choisir un moteur", df_filtered["M_moteur"].dropna().unique())
code_frigo = st.selectbox("Choisir un groupe frigo", df_filtered["C_Groupe frigo"].dropna().unique())
code_hayon = st.selectbox("Choisir un hayon", df_filtered["C_Hayon elevateur"].dropna().unique())

# Fonction pour extraire les détails à partir du code
def get_details_by_code(code):
    if pd.isna(code):
        return "Détails indisponibles"
    rows = df[df.apply(lambda row: code in row.values, axis=1)]
    if rows.empty:
        return "Détails introuvables"
    return str(rows.iloc[0].to_dict())

# Génération du fichier Excel
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiche Technique"

    # Logo
    logo_path = "petit_forestier_logo_officiel.png"
    logo = XLImage(logo_path)
    logo.width = 100
    logo.height = 40
    ws.add_image(logo, "A1")

    # Contenu
    ws.append(["Fiche Technique"])
    ws.append([""])
    ws.append(["Modèle", modele])
    ws.append(["Cabine", code_cabine])
    ws.append(["Détail cabine", get_details_by_code(code_cabine)])
    ws.append(["Châssis", code_chassis])
    ws.append(["Détail châssis", get_details_by_code(code_chassis)])
    ws.append(["Caisse", code_caisse])
    ws.append(["Détail caisse", get_details_by_code(code_caisse)])
    ws.append(["Moteur", code_moteur])
    ws.append(["Détail moteur", get_details_by_code(code_moteur)])
    ws.append(["Groupe Frigo", code_frigo])
    ws.append(["Détail frigo", get_details_by_code(code_frigo)])
    ws.append(["Hayon", code_hayon])
    ws.append(["D
