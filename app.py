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

# Colonnes requises
required_columns = ["Code_Pays", "Marque", "Modele", "Code_PF", "Standard_PF", "C_Cabine", "M_Moteur", "C_Chassis", "C_Caisse", "C_Groupe Frigorifique", "C_Hayon"]
if not all(col in df.columns for col in required_columns):
    st.error("Colonnes manquantes dans le fichier Excel: " + ", ".join(required_columns))
    st.stop()

# --------- Menus déroulants dans le bon ordre ---------

# 1. Code_Pays
code_pays = st.selectbox("Choisir un code pays", sorted(df["Code_Pays"].dropna().unique()))
df_filtered = df[df["Code_Pays"] == code_pays]

# 2. Marque (filtré par code pays)
marque = st.selectbox("Choisir une marque", sorted(df_filtered["Marque"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Marque"] == marque]

# 3. Modèle (filtré par code pays + marque)
modele = st.selectbox("Choisir un modèle", sorted(df_filtered["Modele"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Modele"] == modele]

# 4. Code_PF (filtré par code pays + marque + modèle)
code_pf = st.selectbox("Choisir un Code PF", sorted(df_filtered["Code_PF"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Code_PF"] == code_pf]

# Composants (après tous les filtres)
code_cabine = st.selectbox("Choisir une cabine", df_filtered["C_Cabine"].dropna().unique())
code_chassis = st.selectbox("Choisir un châssis", df_filtered["C_Chassis"].dropna().unique())
code_caisse = st.selectbox("Choisir une caisse", df_filtered["C_Caisse"].dropna().unique())
code_moteur = st.selectbox("Choisir un moteur", df_filtered["M_Moteur"].dropna().unique())
code_frigo = st.selectbox("Choisir un groupe frigo", df_filtered["C_Groupe Frigorifique"].dropna().unique())
code_hayon = st.selectbox("Choisir un hayon", df_filtered["C_Hayon"].dropna().unique())

# --------- Détails par code ---------
def get_details_by_code(code):
    if pd.isna(code):
        return "Détails indisponibles"
    rows = df[df.apply(lambda row: code in row.values, axis=1)]
    if rows.empty:
        return "Détails introuvables"
    return str(rows.iloc[0].to_dict())

# --------- Génération Excel ---------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiche Technique"

    logo_path = "petit_forestier_logo_officiel.png"
    logo = XLImage(logo_path)
    logo.width = 100
    logo.height = 40
    ws.add_image(logo, "A1")

    ws.append(["Fiche Technique"])
    ws.append([""])
    ws.append(["Code Pays", code_pays])
    ws.append(["Marque", marque])
    ws.append(["Modèle", modele])
    ws.append(["Code PF", code_pf])
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
    ws.append(["Détail hayon", get_details_by_code(code_hayon)])

    output = BytesIO()
    wb.save(output)
    return output

# --------- Bouton d'export ---------
if st.button("📄 Générer la fiche technique"):
    excel_file = generate_excel()
    st.success("✅ Fiche technique générée avec succès !")
    st.download_button(label="💾 Télécharger le fichier Excel",
                       data=excel_file.getvalue(),
                       file_name="fiche_technique.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


