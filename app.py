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

# Logo & Titre
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

# Filtres utilisateur
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

# Extraction de données
def get_details(df_component, code, code_column="Code"):
    if code in df_component[code_column].values:
        return df_component[df_component[code_column] == code].iloc[0].to_dict()
    return {}

# Génération de la fiche technique
def generate_filled_ft():
    wb = load_workbook("Modèle FT.xlsx")
    ws = wb.active

    # Infos générales
    ws["E8"] = code_pays
    ws["E9"] = marque
    ws["E10"] = modele
    ws["E11"] = code_pf

    # Cabine
    cabine_data = get_details(cabine_df, code_cabine)
    ws["E15"] = cabine_data.get("Code", "")
    ws["E16"] = cabine_data.get("Marque", "")
    ws["E17"] = cabine_data.get("Modèle", "")
    ws["E18"] = cabine_data.get("Version", "")

    # Châssis
    chassis_data = get_details(chassis_df, code_chassis)
    ws["E21"] = chassis_data.get("Code", "")
    ws["E22"] = chassis_data.get("PTAC", "")
    ws["E23"] = chassis_data.get("Empattement", "")

    # Caisse
    caisse_data = get_details(caisse_df, code_caisse)
    ws["E26"] = caisse_data.get("Code", "")
    ws["E27"] = caisse_data.get("Longueur", "")
    ws["E28"] = caisse_data.get("Largeur", "")

    # Moteur
    moteur_data = get_details(moteur_df, code_moteur)
    ws["E31"] = moteur_data.get("Code", "")
    ws["E32"] = moteur_data.get("Puissance", "")

    # Frigo
    frigo_data = get_details(frigo_df, code_frigo)
    ws["E35"] = frigo_data.get("Code", "")
    ws["E36"] = frigo_data.get("Marque", "")
    ws["E37"] = frigo_data.get("Modèle", "")

    # Hayon
    hayon_data = get_details(hayon_df, code_hayon)
    ws["E40"] = hayon_data.get("Code", "")
    ws["E41"] = hayon_data.get("Capacité", "")

    # Export en mémoire
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Bouton de téléchargement
st.download_button(
    label="📥 Télécharger la fiche technique complète",
    data=generate_filled_ft(),
    file_name=f"FT_{code_pf}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
