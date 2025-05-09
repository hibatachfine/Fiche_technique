import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from PIL import Image

# Authentification par mot de passe
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

# Titre et logo
st.image("petit_forestier_logo_officiel.png", width=700)
st.markdown("<h1 style='color:#057A20;'>Générateur de Fiches Techniques</h1>", unsafe_allow_html=True)
st.markdown("---")

try:
    df = pd.read_excel("bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")
except Exception as e:
    st.error(f"Erreur lors du chargement du fichier Excel : {e}")
    st.stop()

required_columns = ["Code_Pays", "Marque", "Modele", "Code_PF", "Standard_PF", "C_Cabine", "M_Moteur", "C_Chassis", "C_Caisse", "C_Groupe Frigorifique", "C_Hayon"]
if not all(col in df.columns for col in required_columns):
    st.error("Colonnes manquantes dans le fichier Excel: " + ", ".join(required_columns))
    st.stop()

# Sélections filtrées
code_pays = st.selectbox("Choisir un code pays", sorted(df["Code_Pays"].dropna().unique()))
df_filtered = df[df["Code_Pays"] == code_pays]

marque = st.selectbox("Choisir une marque", sorted(df_filtered["Marque"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Marque"] == marque]

modele = st.selectbox("Choisir un modèle", sorted(df_filtered["Modele"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Modele"] == modele]

code_pf = st.selectbox("Choisir un Code PF", sorted(df_filtered["Code_PF"].dropna().unique()))
df_filtered = df_filtered[df_filtered["Code_PF"] == code_pf]

code_cabine = st.selectbox("Choisir une cabine", df_filtered["C_Cabine"].dropna().unique())
code_chassis = st.selectbox("Choisir un châssis", df_filtered["C_Chassis"].dropna().unique())
code_caisse = st.selectbox("Choisir une caisse", df_filtered["C_Caisse"].dropna().unique())
code_moteur = st.selectbox("Choisir un moteur", df_filtered["M_Moteur"].dropna().unique())
code_frigo = st.selectbox("Choisir un groupe frigorifique", df_filtered["C_Groupe Frigorifique"].dropna().unique())
code_hayon = st.selectbox("Choisir un hayon", df_filtered["C_Hayon"].dropna().unique())

# Détails multiples

def get_all_details_by_code(code):
    if pd.isna(code):
        return "Détails indisponibles"
    rows = df[df.apply(lambda row: code in row.values, axis=1)]
    if rows.empty:
        return "Détails introuvables"
    all_matches = []
    for _, row in rows.iterrows():
        all_matches.append(str(row.to_dict()))
    return "\n\n".join(all_matches)

# Excel stylé

def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiche Technique"

    logo_path = "petit_forestier_logo_officiel.png"
    logo = XLImage(logo_path)
    logo.width = 300
    logo.height = 40
    ws.add_image(logo, "A1")

    ws.merge_cells('A3:B3')
    cell = ws['A3']
    cell.value = "FICHE TECHNIQUE"
    cell.font = Font(size=14, bold=True, color="057A20")
    cell.alignment = Alignment(horizontal="center")

    bold_green = Font(bold=True, color="057A20")
    wrap = Alignment(wrap_text=True, vertical="top")
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    fill_gray = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")

    rows = [
        ("Code Pays", code_pays),
        ("Marque", marque),
        ("Modèle", modele),
        ("Code PF", code_pf),
        ("Cabine", code_cabine),
        ("Détails cabine", get_all_details_by_code(code_cabine)),
        ("Châssis", code_chassis),
        ("Détails châssis", get_all_details_by_code(code_chassis)),
        ("Caisse", code_caisse),
        ("Détails caisse", get_all_details_by_code(code_caisse)),
        ("Moteur", code_moteur),
        ("Détails moteur", get_all_details_by_code(code_moteur)),
        ("Groupe Frigo", code_frigo),
        ("Détails frigo", get_all_details_by_code(code_frigo)),
        ("Hayon", code_hayon),
        ("Détails hayon", get_all_details_by_code(code_hayon)),
    ]

    row_num = 5
    for label, value in rows:
        ws.cell(row=row_num, column=1, value=label).font = bold_green
        ws.cell(row=row_num, column=2, value=value)
        ws.cell(row=row_num, column=1).alignment = wrap
        ws.cell(row=row_num, column=2).alignment = wrap
        ws.cell(row=row_num, column=1).border = border
        ws.cell(row=row_num, column=2).border = border
        if row_num % 2 == 0:
            ws.cell(row=row_num, column=1).fill = fill_gray
            ws.cell(row=row_num, column=2).fill = fill_gray
        row_num += 1

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 80

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Export bouton
st.download_button(
    label="📀 Télécharger la fiche technique",
    data=generate_excel().getvalue(),
    file_name="fiche_technique.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
) 
