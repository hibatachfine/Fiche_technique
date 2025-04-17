import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook, load_workbook
from PIL import Image

# Charger logo
st.set_page_config(layout="centered")
logo = Image.open("petit_forestier_logo_officiel.png")
st.image(logo, width=180)

# Titre principal
st.markdown("<h1 style='text-align: center; color: #017a0c;'>Générateur de Fiche Technique</h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# Charger les données
df = pd.read_excel("bdd_ht.xlsx", sheet_name=0)

# Nettoyer les noms de colonnes
df.columns = df.columns.str.strip().str.upper()

# Vérifier que les colonnes essentielles existent
required_columns = ["MODELE", "CABINE", "CHASSIS", "CAISSE", "MOTEUR", "FRIGO", "HAYON"]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    st.error(f"Colonnes manquantes dans le fichier Excel: {', '.join(missing_columns)}")
    st.stop()

# Liste des modèles disponibles
modele_list = sorted(df["MODELE"].dropna().unique())
modele = st.selectbox("🛠️ Choisir un modèle", modele_list)

# Filtrer les lignes qui correspondent au modèle choisi
filtered_df = df[df["MODELE"] == modele]

# Si plusieurs lignes existent, prendre la première
selected_row = filtered_df.iloc[0]

# Menus déroulants basés sur le modèle sélectionné
code_cabine = st.selectbox("🚐 Choisir une cabine", filtered_df["CABINE"].dropna().unique())
code_chassis = st.selectbox("🛞 Choisir un châssis", filtered_df["CHASSIS"].dropna().unique())
code_caisse = st.selectbox("🚚 Choisir une caisse", filtered_df["CAISSE"].dropna().unique())
code_moteur = st.selectbox("⚙️ Choisir un moteur", filtered_df["MOTEUR"].dropna().unique())
code_frigo = st.selectbox("❄️ Choisir un groupe frigo", filtered_df["FRIGO"].dropna().unique())
code_hayon = st.selectbox("🔧 Choisir un hayon", filtered_df["HAYON"].dropna().unique() if filtered_df["HAYON"].notna().any() else ["Aucun"])

# Génération du fichier Excel
def generate_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiche Technique"

    ws.append(["Élément", "Code sélectionné"])
    for key, value in data.items():
        ws.append([key, value])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# Bouton de génération
if st.button("📄 Générer la fiche technique"):
    data = {
        "Modèle": modele,
        "Cabine": code_cabine,
        "Châssis": code_chassis,
        "Caisse": code_caisse,
        "Moteur": code_moteur,
        "Frigo": code_frigo,
        "Hayon": code_hayon,
    }
    excel_file = generate_excel(data)
    st.success("✅ Fiche technique générée avec succès !")
    st.download_button("⬇️ Télécharger la fiche technique", data=excel_file, file_name="fiche_technique.xlsx")
