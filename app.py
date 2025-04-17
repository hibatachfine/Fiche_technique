import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

# Charger les données
df_main = pd.read_excel("bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")
df_cabine = pd.read_excel("bdd_ht.xlsx", sheet_name="CABINES")
df_chassis = pd.read_excel("bdd_ht.xlsx", sheet_name="CHASSIS")
df_moteur = pd.read_excel("bdd_ht.xlsx", sheet_name="MOTEURS")
df_frigo = pd.read_excel("bdd_ht.xlsx", sheet_name="FRIGO")
df_caisse = pd.read_excel("bdd_ht.xlsx", sheet_name="CAISSES")
df_hayon = pd.read_excel("bdd_ht.xlsx", sheet_name="HAYONS")

st.title("Générateur de Fiche Technique")

modeles = df_main["Modele"].dropna().unique()
modele_select = st.selectbox("Choisir un modèle", modeles)

ligne = df_main[df_main["Modele"] == modele_select].iloc[0]

code_cabine = ligne["C_Cabine"]
code_chassis = ligne["C_Chassis"]
code_caisse = ligne["C_Caisse"]
code_moteur = ligne["M_moteur"]
code_frigo = ligne["C_Groupe frigo"]
code_hayon = ligne["C_Hayon elevateur"]

st.markdown("### Éléments sélectionnés :")
st.write(f"Cabine: {code_cabine}")
st.write(f"Châssis: {code_chassis}")
st.write(f"Caisse: {code_caisse}")
st.write(f"Moteur: {code_moteur}")
st.write(f"Frigo: {code_frigo}")
st.write(f"Hayon: {code_hayon}")

if st.button("Générer la fiche technique"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiche Technique"

    ws["A1"] = "Modèle"
    ws["B1"] = modele_select

    def write_details(df, code, title, start_row):
        ws[f"A{start_row}"] = title
        bloc = df[df[df.columns[0]] == code]
        for i, col in enumerate(bloc.columns):
            ws.cell(row=start_row + 1, column=i + 1, value=col)
            ws.cell(row=start_row + 2, column=i + 1, value=str(bloc[col].values[0]))
        return start_row + 4

    row = 3
    row = write_details(df_cabine, code_cabine, "Cabine", row)
    row = write_details(df_chassis, code_chassis, "Châssis", row)
    row = write_details(df_caisse, code_caisse, "Caisse", row)
    row = write_details(df_moteur, code_moteur, "Moteur", row)
    row = write_details(df_frigo, code_frigo, "Frigo", row)
    row = write_details(df_hayon, code_hayon, "Hayon", row)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success("Fiche technique générée avec succès !")
    st.download_button(
        label="📄 Télécharger la fiche technique",
        data=buffer,
        file_name=f"Fiche_{modele_select}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )