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

# --- Filtres utilisateur ---
code_pays = st.selectbox("Code pays", sorted(df["Code_Pays"].dropna().unique()))
df_filtered = df[df["Code_Pays"] == code_pays]

marque = st.selectbox("Marque", sorted(df_filtered["Marque"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Marque"] == marque]

modele = st.selectbox("Modèle", sorted(df_filtered["Modele"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Modele"] == modele]

code_pf = st.selectbox("Code PF", sorted(df_filtered["Code_PF"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Code_PF"] == code_pf]

if "Standard_PF" in df_filtered.columns and not df_filtered["Standard_PF"].dropna().empty:
    standard_pf = st.selectbox("Standard PF", sorted(df_filtered["Standard_PF"].dropna().unique()))
    df_filtered = df_filtered[df_filtered["Standard_PF"] == standard_pf]
else:
    standard_pf = ""
    st.warning("Aucune valeur Standard PF trouvée pour ce Code PF.")

code_cabine = st.selectbox("Cabine", df_filtered["C_Cabine"].dropna().unique())
code_chassis = st.selectbox("Châssis", df_filtered["C_Chassis"].dropna().unique())
code_caisse = st.selectbox("Caisse", df_filtered["C_Caisse"].dropna().unique())
code_moteur = st.selectbox("Moteur", df_filtered["M_Moteur"].dropna().unique())
code_frigo = st.selectbox("Groupe Frigorifique", df_filtered["C_Groupe Frigorifique"].dropna().unique())
code_hayon = st.selectbox("Hayon", df_filtered["C_Hayon"].dropna().unique())

# --- Fonctions utilitaires ---
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
        try:
            value = str(crit).strip() if crit is not None else ""
            ws[f"{col_letter}{start_row + i}"] = value
        except Exception as e:
            print(f"Erreur cellule {col_letter}{start_row + i} : {e}")

def generate_filled_ft():
    wb = load_workbook("Modèle FT.xlsx")
    ws = wb["TYPE_FROID"]  # 🟢 Cible directement la bonne feuille

    # Récupération de la ligne sélectionnée
    selected_row = df_filtered.iloc[0]

    # --- Dimensions principales (bloc en haut à droite)
    ws["J6"] = selected_row.get("L", "")
    ws["J7"] = selected_row.get("Z", "")
    ws["F6"] = selected_row.get("W int utile sur plinthe", "")
    ws["F7"] = selected_row.get("L int utile sur plinthe", "")
    ws["F8"] = selected_row.get("H int", "")
    ws["J8"] = selected_row.get("Hc", "")
    ws["J9"] = selected_row.get("F", "")
    ws["J10"] = selected_row.get("X", "")

    # --- Bloc PTAC
    ws["H15"] = selected_row.get("PTAC", "")
    ws["H16"] = selected_row.get("CU", "")
    ws["H17"] = selected_row.get("Volume", "")
    ws["H18"] = selected_row.get("palettes 800 x 1200 mm", "")

    # --- Infos générales
    ws["B4"] = marque
    ws["C4"] = modele
    ws["E4"] = code_pf
    ws["G4"] = standard_pf

    # --- Insertion critère sous critère (composants)
    insert_criteria(ws, "B22", get_criteria_list(cabine_df, code_cabine, "C_Cabine"))
    insert_criteria(ws, "E22", get_criteria_list(moteur_df, code_moteur, "M_moteur"))
    insert_criteria(ws, "G22", get_criteria_list(chassis_df, code_chassis, "C_Chassis"))
    insert_criteria(ws, "B54", get_criteria_list(caisse_df, code_caisse, "C_Caisse"))
    insert_criteria(ws, "B64", get_criteria_list(frigo_df, code_frigo, "C_Groupe Frigorifique"))
    insert_criteria(ws, "B73", get_criteria_list(hayon_df, code_hayon, "C_Hayon"))

    # Export
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- Téléchargement ---
st.download_button(
    label="Télécharger la fiche technique",
    data=generate_filled_ft(),
    file_name=f"FT_{code_pf}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
