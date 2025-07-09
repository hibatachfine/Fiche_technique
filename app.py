import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

# --- Authentification ---
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

# --- Interface ---
st.image("petit_forestier_logo_officiel.png", width=700)
st.markdown("<h1 style='color:#057A20;'>Générateur de Fiches Techniques</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- Chargement des fichiers Excel ---
try:
    df = pd.read_excel("bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")
    cabine_df = pd.read_excel("bdd_ht.xlsx", sheet_name="CABINES")
    chassis_df = pd.read_excel("bdd_ht.xlsx", sheet_name="CHASSIS")
    caisse_df = pd.read_excel("bdd_ht.xlsx", sheet_name="CAISSES")
    moteur_df = pd.read_excel("bdd_ht.xlsx", sheet_name="MOTEURS")
    frigo_df = pd.read_excel("bdd_ht.xlsx", sheet_name="FRIGO")
    hayon_df = pd.read_excel("bdd_ht.xlsx", sheet_name="HAYONS")
except Exception as e:
    st.error(f"Erreur lors du chargement des fichiers : {e}")
    st.stop()

# --- Sélection des filtres ---
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

# --- Fonction pour extraire les données d'une feuille ---
def get_details(df_component, code, code_column="Code"):
    if code in df_component[code_column].values:
        return df_component[df_component[code_column] == code].iloc[0].to_dict()
    return {}

# --- Génération de la fiche technique Excel ---
def generate_filled_ft():
    wb = load_workbook("Modèle FT.xlsx")
    ws = wb.active

    # Informations générales
    ws["E8"] = code_pays
    ws["E9"] = marque
    ws["E10"] = modele
    ws["E11"] = code_pf

    # --- Cabine ---
    cabine_data = get_details(cabine_df, code_cabine, code_column="C_Cabine")
    ws["E15"] = cabine_data.get("C_Cabine", "")
    ws["E16"] = cabine_data.get("Marque", "")
    ws["E17"] = cabine_data.get("Modèle", "")
    ws["E18"] = cabine_data.get("Version", "")

    # --- Châssis ---
    chassis_data = get_details(chassis_df, code_chassis, code_column="C_Chassis")
    ws["E21"] = chassis_data.get("C_Chassis", "")
    ws["E22"] = chassis_data.get("PTAC", "")
    ws["E23"] = chassis_data.get("Empattement", "")

    # --- Caisse ---
    caisse_data = get_details(caisse_df, code_caisse, code_column="C_Caisse")
    ws["E26"] = caisse_data.get("C_Caisse", "")
    ws["E27"] = caisse_data.get("Longueur", "")
    ws["E28"] = caisse_data.get("Largeur", "")

    # --- Moteur ---
    moteur_data = get_details(moteur_df, code_moteur, code_column="M_moteur")
    ws["E31"] = moteur_data.get("M_moteur", "")
    ws["E32"] = moteur_data.get("Puissance", "")

    # --- Groupe Frigorifique ---
    frigo_data = get_details(frigo_df, code_frigo, code_column="C_Groupe Frigorifique")
    ws["E35"] = frigo_data.get("C_Groupe Frigorifique", "")
    ws["E36"] = frigo_data.get("Marque groupe", "")
    ws["E37"] = frigo_data.get("Modèle groupe", "")

    # --- Hayon ---
    hayon_data = get_details(hayon_df, code_hayon, code_column="C_Hayon")
    ws["E40"] = hayon_data.get("C_Hayon", "")
    ws["E41"] = hayon_data.get("Capacité", "") or hayon_data.get("Puissance", "")

    # --- Export ---
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Bouton de téléchargement ---
st.download_button(
    label="📥 Télécharger la fiche technique complète",
    data=generate_filled_ft(),
    file_name=f"FT_{code_pf}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
