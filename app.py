
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

# Configuration de la page
st.set_page_config(page_title="Fiche Technique", layout="centered")

# Affichage du logo centré
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image("petit_forestier_logo_officiel.png", width=200)

# Titre principal
st.markdown("<h1 style='text-align: center; color: #053B06;'>Générateur de Fiche Technique</h1>", unsafe_allow_html=True)
st.markdown("---")

# Chargement de la base de données
df_modeles = pd.read_excel("bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")
df_cabines = pd.read_excel("bdd_ht.xlsx", sheet_name="CABINES")
df_caisses = pd.read_excel("bdd_ht.xlsx", sheet_name="CAISSES")
df_chassis = pd.read_excel("bdd_ht.xlsx", sheet_name="CHASSIS")
df_frigo = pd.read_excel("bdd_ht.xlsx", sheet_name="FRIGO")
df_hayon = pd.read_excel("bdd_ht.xlsx", sheet_name="HAYONS")
df_moteurs = pd.read_excel("bdd_ht.xlsx", sheet_name="MOTEURS")

# Sélection du modèle
modele = st.selectbox("🛠️ Choisir un modèle", sorted(df_modeles["MODELE"].dropna().unique()))

# Filtrage des composants selon le modèle
bloc = df_modeles[df_modeles["MODELE"] == modele].iloc[0]

code_cabine = bloc["CABINE"]
code_chassis = bloc["CHASSIS"]
code_caisse = bloc["CAISSE"]
code_moteur = bloc["MOTEUR"]
code_frigo = bloc["FRIGO"]
code_hayon = bloc["HAYON"]

# Fonctions de recherche des lignes correspondantes
def get_row(df, code):
    return df[df["CODE"] == code]

# Création des sélections dynamiques
col_cab, col_chas = st.columns(2)
with col_cab:
    st.selectbox("🚖 Choisir une cabine", df_cabines["CODE"], index=df_cabines[df_cabines["CODE"] == code_cabine].index[0])
with col_chas:
    st.selectbox("🦾 Choisir un châssis", df_chassis["CODE"], index=df_chassis[df_chassis["CODE"] == code_chassis].index[0])

col_cai, col_mot = st.columns(2)
with col_cai:
    st.selectbox("📦 Choisir une caisse", df_caisses["CODE"], index=df_caisses[df_caisses["CODE"] == code_caisse].index[0])
with col_mot:
    st.selectbox("🔧 Choisir un moteur", df_moteurs["CODE"], index=df_moteurs[df_moteurs["CODE"] == code_moteur].index[0])

col_fri, col_hay = st.columns(2)
with col_fri:
    st.selectbox("❄️ Choisir un groupe frigo", df_frigo["CODE"], index=df_frigo[df_frigo["CODE"] == code_frigo].index[0])
with col_hay:
    hayon_index = df_hayon[df_hayon["CODE"] == code_hayon].index
    st.selectbox("⛓️ Choisir un hayon", df_hayon["CODE"], index=hayon_index[0] if not hayon_index.empty else 0)

# Espace
st.markdown("")

# Génération de la fiche technique
def write_details(df, code, nom_bloc, ws, start_row):
    bloc = df[df["CODE"] == code]
    if bloc.empty:
        return
    ws.cell(row=start_row, column=1, value=nom_bloc)
    for col in bloc.columns:
        ws.cell(row=start_row+1, column=list(bloc.columns).index(col)+1, value=str(bloc[col].values[0]))

def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiche Technique"

    write_details(df_cabines, code_cabine, "Cabine", ws, 1)
    write_details(df_chassis, code_chassis, "Châssis", ws, 5)
    write_details(df_caisses, code_caisse, "Caisse", ws, 9)
    write_details(df_moteurs, code_moteur, "Moteur", ws, 13)
    write_details(df_frigo, code_frigo, "Frigo", ws, 17)
    write_details(df_hayon, code_hayon, "Hayon", ws, 21)

    output = BytesIO()
    wb.save(output)
    return output

if st.button("📄 Générer la fiche technique"):
    result = generate_excel()
    st.success("Fiche technique générée ✅")
    st.download_button(label="⬇️ Télécharger", data=result.getvalue(), file_name="fiche_technique.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
