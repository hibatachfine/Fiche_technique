import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from shutil import copyfile

# Charger la base principale
df_ref = pd.read_excel("/mnt/data/bdd_ht.xlsx", sheet_name="FS_referentiel_produits_std")

# Lire chaque composant avec leur colonne de code respective
df_cabines = pd.read_excel("/mnt/data/bdd_ht.xlsx", sheet_name="CABINES")
df_chassis = pd.read_excel("/mnt/data/bdd_ht.xlsx", sheet_name="CHASSIS")
df_caisses = pd.read_excel("/mnt/data/bdd_ht.xlsx", sheet_name="CAISSES")
df_frigo = pd.read_excel("/mnt/data/bdd_ht.xlsx", sheet_name="FRIGO")
df_hayon = pd.read_excel("/mnt/data/bdd_ht.xlsx", sheet_name="HAYONS")
df_moteur = pd.read_excel("/mnt/data/bdd_ht.xlsx", sheet_name="MOTEURS")

# Ligne de test (ex: ligne 0)
row = df_ref.iloc[0]
code_pf = row["Code_PF"]
code_pays = row["Code_Pays"]
marque = row["Marque"]
modele = row["Modele"]

# Fonction pour récupérer les détails à partir du nom de colonne code
def get_details(code, df, code_col_name):
    if code in df[code_col_name].values:
        data = df[df[code_col_name] == code].iloc[0].drop(code_col_name)
        return "\n".join([f"{col}: {val}" for col, val in data.items() if pd.notna(val)])
    return "Non trouvé"

# Récupération des détails avec les bons noms de colonnes
details_cabine = get_details(row["C_Cabine"], df_cabines, "Code_Cabine")
details_chassis = get_details(row["C_Chassis"], df_chassis, "Code_Chassis")
details_caisse = get_details(row["C_Caisse"], df_caisses, "Code_Caisse")
details_frigo = get_details(row["C_Groupe Frigorifique"], df_frigo, "Code_Frigorifique")
details_hayon = get_details(row["C_Hayon"], df_hayon, "Code_Hayon")
details_moteur = get_details(row["M_Moteur"], df_moteur, "Code_Moteur")

# Copier le modèle
template_path = "/mnt/data/Modèle FT.xlsx"
output_path = "/mnt/data/fiche_technique_remplie.xlsx"
copyfile(template_path, output_path)

# Remplissage du fichier
wb = load_workbook(output_path)
ws = wb.active

ws["E8"] = code_pf
ws["E9"] = code_pays
ws["E10"] = marque
ws["E11"] = modele

ws["E15"] = row["C_Cabine"]
ws["F15"] = details_cabine

ws["E16"] = row["C_Chassis"]
ws["F16"] = details_chassis

ws["E17"] = row["C_Caisse"]
ws["F17"] = details_caisse

ws["E18"] = row["C_Groupe Frigorifique"]
ws["F18"] = details_frigo

ws["E19"] = row["C_Hayon"]
ws["F19"] = details_hayon

ws["E20"] = row["M_Moteur"]
ws["F20"] = details_moteur

# Texte aligné haut et retour à la ligne
for row in ws.iter_rows(min_row=15, max_row=20, min_col=6, max_col=6):
    for cell in row:
        cell.alignment = Alignment(wrap_text=True, vertical="top")

wb.save(output_path)
print("✅ Fiche technique remplie avec succès :", output_path)
