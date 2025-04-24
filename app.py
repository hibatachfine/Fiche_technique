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

# 3. Modèle (filtré par code pays/marque)
modele = st.selectbox("Choisir un modèle", sorted(df_filtered["Modele"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Modele"] == modele]

# 4. Code_PF (filtré par code pays/marque/modèle)
code_pf = st.selectbox("Choisir un Code PF", sorted(df_filtered["Code_PF"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Code_PF"] == code_pf]

# Composants (après tous les filtres)
code_cabine = st.selectbox("Choisir une cabine", df_filtered["C_Cabine"].dropna().unique())
code_chassis = st.selectbox("Choisir un châssis", df_filtered["C_Chassis"].dropna().unique())
code_caisse = st.selectbox("Choisir une caisse", df_filtered["C_Caisse"].dropna().unique())
code_moteur = st.selectbox("Choisir un moteur", df_filtered["M_Moteur"].dropna().unique())
code_frigo = st.selectbox("Choisir un groupe frigorifique", df_filtered["C_Groupe Frigorifique"].dropna().unique())
code_hayon = st.selectbox("Choisir un hayon", df_filtered["C_Hayon"].dropna().unique())

# --------- Fonction pour récupérer les détails depuis les fichiers externes ---------
def get_details_from_file(file_name, code):
    try:
        details_df = pd.read_excel(file_name)
        details = details_df[details_df["Code"] == code]
        if details.empty:
            return "Détails non trouvés"
        return str(details.iloc[0].to_dict())
    except Exception as e:
        return f"Erreur lors du chargement des détails : {e}"

# --------- Génération de l'Excel ---------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiche Technique"

    logo_path = "petit_forestier_logo_officiel.png"
    logo = XLImage(logo_path)
    logo.width = 300
    logo.height = 40
    ws.add_image(logo, "A1")

    ws.append(["Fiche Technique"])
    ws.append([""])
    ws.append(["Code Pays", code_pays])
    ws.append(["Marque", marque])
    ws.append(["Modèle", modele])
    ws.append(["Code PF", code_pf])

    # Ajouter les détails pour chaque élément
    ws.append(["Cabine", C_Cabine])
    ws.append(["Détails cabine", get_details_from_file("CABINES.xlsx", C_Cabine)])

    ws.append(["Châssis", C_Chassis])
    ws.append(["Détails châssis", get_details_from_file("CHASSIS.xlsx", C_Chassis)])

    ws.append(["Caisse", C_Caisse])
    ws.append(["Détails caisse", get_details_from_file("CAISSES.xlsx", C_Caisse)])

    ws.append(["Moteur", M_moteur])
    ws.append(["Détails moteur", get_details_from_file("MOTEURS.xlsx", M_moteur)])

    ws.append(["Groupe Frigo", C_Groupe Frigorifique])
    ws.append(["Détails frigo", get_details_from_file("FRIGO.xlsx", C_Groupe Frigorifique)])

    ws.append(["Hayon", C_Hayon])
    ws.append(["Détails hayon", get_details_from_file("HAYONS.xlsx", C_Hayon)])

    output = BytesIO()
    wb.save(output)
    return output

# --------- Bouton d'export ---------
st.download_button(label="💾 Télécharger la fiche technique",
                   data=generate_excel().getvalue(),
                   file_name="fiche_technique.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
