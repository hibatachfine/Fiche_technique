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

# --- Normalisation des noms de colonnes ---
df.columns = df.columns.str.replace('\n', ' ').str.strip()

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
    col = ''.join(filter(str.isalpha, start_cell))
    try:
        row = int(''.join(filter(str.isdigit, start_cell)))
    except ValueError:
        return

    for i, item in enumerate(criteria_list):
        try:
            cell_ref = f"{col}{row + i}"
            value = str(item).strip() if item is not None else ""
            ws[cell_ref] = value
        except Exception as e:
            print(f"Erreur cellule {cell_ref} : {e}")

def insert_criteria_dual_column(ws, start_cell_1, limit_row_1, start_cell_2, criteria_list):
    col1 = ''.join(filter(str.isalpha, start_cell_1))
    row1 = int(''.join(filter(str.isdigit, start_cell_1)))
    end_row1 = limit_row_1

    col2 = ''.join(filter(str.isalpha, start_cell_2))
    row2 = int(''.join(filter(str.isdigit, start_cell_2)))

    for i, item in enumerate(criteria_list):
        value = str(item).strip()
        if row1 + i <= end_row1:
            cell_ref = f"{col1}{row1 + i}"
        else:
            cell_ref = f"{col2}{row2 + (i - (end_row1 - row1 + 1))}"
        ws[cell_ref] = value

def safe_excel_set(ws, cell, value, label=""):
    try:
        # Vérifier si la cellule fait partie d'un bloc fusionné
        for merged_range in ws.merged_cells.ranges:
            if cell in merged_range:
                cell = merged_range.coord.split(":")[0]  # première cellule fusionnée
                break

        # Écrire la valeur si elle est définie
        ws[cell] = str(value) if pd.notna(value) else ""
    except Exception as e:
        st.error(f"Erreur cellule {cell} ({label}) : {e}")


# --- Génération de la fiche technique ---
def generate_filled_ft():
    wb = load_workbook("Modèle FT.xlsx")

    if "TYPE_FROID" not in wb.sheetnames:
        st.error("La feuille 'TYPE_FROID' est introuvable dans le fichier Excel.")
        st.stop()

    ws = wb["TYPE_FROID"]

    matching_rows = df[df["Code_PF"] == code_pf]
    if matching_rows.empty:
        st.error("Aucune ligne correspondante trouvée pour le Code PF sélectionné.")
        st.stop()

    selected_row = matching_rows.iloc[0]

    # Dimensions
    safe_excel_set(ws, "J6", selected_row.get("L", ""), "L")
    safe_excel_set(ws, "J7", selected_row.get("Z", ""), "Z")
    safe_excel_set(ws, "J8", selected_row.get("Hc", ""), "Hc")
    safe_excel_set(ws, "J9", selected_row.get("F", ""), "F")
    safe_excel_set(ws, "J10", selected_row.get("X", ""), "X")
    safe_excel_set(ws, "H7", selected_row.get("W int utile sur plinthe", ""), "W utile")
    safe_excel_set(ws, "H8", selected_row.get("L int utile sur plinthe", ""), "L utile")
    safe_excel_set(ws, "H9", selected_row.get("H", ""), "H")

    # Bloc PTAC
    safe_excel_set(ws, "H12", selected_row.get("PTAC", ""), "PTAC")
    safe_excel_set(ws, "H13", selected_row.get("CU", ""), "CU")
    safe_excel_set(ws, "H14", selected_row.get("Volume", ""), "Volume")
    safe_excel_set(ws, "H15", selected_row.get("palettes 800 x 1200 mm", ""), "Palettes")

    # Infos générales
    ws["B2"] = marque
    ws["C2"] = modele
    ws["E2"] = code_pf
    ws["G2"] = standard_pf

    # Insertion critères
    insert_criteria(ws, "B19", get_criteria_list(cabine_df, code_cabine, "C_Cabine"))
    insert_criteria(ws, "E19", get_criteria_list(moteur_df, code_moteur, "M_moteur"))
    insert_criteria(ws, "G19", get_criteria_list(chassis_df, code_chassis, "C_Chassis"))
    insert_criteria(ws, "B38", get_criteria_list(caisse_df, code_caisse, "C_Caisse"))
    insert_criteria_dual_column(ws, "B59", 65, "E59", get_criteria_list(frigo_df, code_frigo, "C_Groupe Frigorifique"))
    insert_criteria_dual_column(ws, "B68", 74, "E68", get_criteria_list(hayon_df, code_hayon, "C_Hayon"))

    # Export fichier
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Téléchargement ---
st.download_button(
    label="📅 Télécharger la fiche technique",
    data=generate_filled_ft(),
    file_name=f"FT_{code_pf}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
