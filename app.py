def generate_filled_ft():
    wb = load_workbook("Modèle FT.xlsx")
    sheet = wb["TYPE_FROID"]  # Feuille cible

    # ✅ Sélection de la ligne filtrée avec les bons critères
    selected_row = df_filtered.iloc[0]

    # --- 📏 Dimensions ---
    sheet["J6"] = selected_row["L"]             # L
    sheet["J7"] = selected_row["Z"]             # Z
    sheet["F6"] = selected_row["W int utile sur plinthe"]
    sheet["F7"] = selected_row["L int utile sur plinthe"]
    sheet["F8"] = selected_row["H"]             # H intérieure
    sheet["J8"] = selected_row["Hc"]
    sheet["J9"] = selected_row["F"]
    sheet["J10"] = selected_row["X"]
    sheet["F11"] = selected_row["palettes 800 x 1200 mm"]

    # --- 📦 PTAC / CU / Volume ---
    sheet["H15"] = selected_row.get("PTAC", "")
    sheet["H16"] = selected_row.get("CU", "")
    sheet["H17"] = selected_row.get("Volume", "")
    sheet["H18"] = selected_row.get("palettes 800 x 1200 mm", "")

    # --- Infos générales ---
    sheet["B4"] = marque
    sheet["C4"] = modele
    sheet["E4"] = code_pf
    sheet["G4"] = standard_pf

    # --- 🧩 Composants ---
    insert_criteria(sheet, "B22", get_criteria_list(cabine_df, code_cabine, "C_Cabine"))
    insert_criteria(sheet, "E22", get_criteria_list(moteur_df, code_moteur, "M_moteur"))
    insert_criteria(sheet, "G22", get_criteria_list(chassis_df, code_chassis, "C_Chassis"))
    insert_criteria(sheet, "B54", get_criteria_list(caisse_df, code_caisse, "C_Caisse"))
    insert_criteria(sheet, "B64", get_criteria_list(frigo_df, code_frigo, "C_Groupe Frigorifique"))
    insert_criteria(sheet, "B73", get_criteria_list(hayon_df, code_hayon, "C_Hayon"))

    # --- Export ---
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
