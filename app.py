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

# Chargement des fichiers
df = pd.read_excel("bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")
cabine_df = pd.read_excel("bdd_ht.xlsx", sheet_name="Cabine")
chassis_df = pd.read_excel("bdd_ht.xlsx", sheet_name="Châssis")
caisse_df = pd.read_excel("bdd_ht.xlsx", sheet_name="Caisse")
moteur_df = pd.read_excel("bdd_ht.xlsx", sheet_name="Moteur")
frigo_df = pd.read_excel("bdd_ht.xlsx", sheet_name="Groupe frigorifique")
hayon_df = pd.read_excel("bdd_ht.xlsx", sheet_name="Hayon")

# Sélections
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

# Fonction pour obtenir les détails du fichier de référence
def get_details(df_component, code, code_column="Code"):
    if code in df_component[code_column].values:
        row = df_component[df_component[code_column] == code].iloc[0]
        print(f"Détails trouvés pour {code}: {row.to_dict()}")  # Débug
        return row.to_dict()
    else:
        print(f"Aucun détail trouvé pour {code}")  # Débug
        return {}

# Générer fichier basé sur modèle
def generate_filled_ft():
    wb = load_workbook("Modèle FT.xlsx")
    ws = wb.active

    # Renseignements simples
    ws["E8"] = code_pays
    ws["E9"] = marque
    ws["E10"] = modele
    ws["E11"] = code_pf

    # Cabine
    cabine_data = get_details(cabine_df, code_cabine)
    if cabine_data:
        ws["E15"] = cabine_data.get("Code", "Non spécifié")
        ws["E16"] = cabine_data.get("Marque", "Non spécifié")
        ws["E17"] = cabine_data.get("Modèle", "Non spécifié")
        ws["E18"] = cabine_data.get("Version", "Non spécifié")
    else:
        ws["E15:E18"] = "Données manquantes"

    # Châssis
    chassis_data = get_details(chassis_df, code_chassis)
    if chassis_data:
        ws["E21"] = chassis_data.get("Code", "Non spécifié")
        ws["E22"] = chassis_data.get("PTAC", "Non spécifié")
        ws["E23"] = chassis_data.get("Empattement", "Non spécifié")
    else:
        ws["E21:E23"] = "Données manquantes"

    # Caisse
    caisse_data = get_details(caisse_df, code_caisse)
    if caisse_data:
        ws["E26"] = caisse_data.get("Code", "Non spécifié")
        ws["E27"] = caisse_data.get("Longueur", "Non spécifié")
        ws["E28"] = caisse_data.get("Largeur", "Non spécifié")
    else:
        ws["E26:E28"] = "Données manquantes"

    # Moteur
    moteur_data = get_details(moteur_df, code_moteur)
    if moteur_data:
        ws["E31"] = moteur_data.get("Code", "Non spécifié")
        ws["E32"] = moteur_data.get("Puissance", "Non spécifié")
    else:
        ws["E31:E32"] = "Données manquantes"

    # Frigo
    frigo_data = get_details(frigo_df, code_frigo)
    if frigo_data:
        ws["E35"] = frigo_data.get("Code", "Non spécifié")
        ws["E36"] = frigo_data.get("Marque", "Non spécifié")
        ws["E37"] = frigo_data.get("Modèle", "Non spécifié")
    else:
        ws["E35:E37"] = "Données manquantes"

    # Hayon
    hayon_data = get_details(hayon_df, code_hayon)
    if hayon_data:
        ws["E40"] = hayon_data.get("Code", "Non spécifié")
        ws["E41"] = hayon_data.get("Capacité", "Non spécifié")
    else:
        ws["E40:E41"] = "Données manquantes"

    # Export
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Bouton téléchargement
st.download_button(
    label="📥 Télécharger la fiche technique complète",
    data=generate_filled_ft(),
    file_name=f"FT_{code_pf}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
