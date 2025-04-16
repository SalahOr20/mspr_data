import pandas as pd

# 1. Charger le fichier avec les deux premières lignes comme en-tête
df_raw = pd.read_excel("../Data/data_population.xlsx", sheet_name="Data_population", header=[0, 1])

# 2. Fusionner les deux lignes d'en-tête
df_raw.columns = [f"{col[1]}" if "Unnamed" not in col[1] else f"{col[0]}" for col in df_raw.columns]

# 3. Nettoyer les lignes invalides (sans code INSEE) et doublons
df_raw = df_raw[df_raw["Code géographique"].notna() & df_raw["Code géographique"].astype(str).str.isdigit()]
df_raw = df_raw.drop_duplicates()

# 4. Sélectionner les colonnes utiles
info_cols = ["Code géographique", "Région", "Département", "Libellé géographique"]
pop_cols = [col for col in df_raw.columns if col.startswith("Population en") and int(col[-4:]) >= 1999]
df = df_raw[info_cols + pop_cols].copy()

# 5. Renommer les colonnes
df.rename(columns={
    "Code géographique": "Code Commune",
    "Région": "Code Région",
    "Département": "Code Département",
    "Libellé géographique": "Commune",
}, inplace=True)

df.rename(columns={col: f"Pop{col[-4:]}" for col in pop_cols}, inplace=True)

# 6. Convertir les types de colonnes
df["Code Commune"] = df["Code Commune"].astype(str)
df["Code Département"] = df["Code Département"].astype(str).str.zfill(2)
df["Code Région"] = df["Code Région"].astype(str)
df["Commune"] = df["Commune"].astype(str)

for col in df.columns:
    if col.startswith("Pop"):
        df[col] = pd.to_numeric(df[col], errors="coerce")

# 7. Filtrer uniquement l’Île-de-France
idf_departements = ['75', '77', '78', '91', '92', '93', '94', '95']
df = df[df["Code Département"].isin(idf_departements)]

# 8. Calculer les indicateurs statistiques
df["Evolution_1999_2022"] = df["Pop2022"] - df["Pop1999"]
df["Taux_Croissance_1999_2022"] = ((df["Pop2022"] - df["Pop1999"]) / df["Pop1999"]) * 100
df["Population_Moy_5Derniers"] = df[["Pop2022", "Pop2021", "Pop2020", "Pop2019", "Pop2018"]].mean(axis=1)

# 🔁 Types explicites
df["Evolution_1999_2022"] = df["Evolution_1999_2022"].fillna(0).astype(int)
df["Taux_Croissance_1999_2022"] = df["Taux_Croissance_1999_2022"].astype(float)
df["Population_Moy_5Derniers"] = df["Population_Moy_5Derniers"].astype(float)

# 9. Sauvegarde en CSV
df.to_csv("../output_data/Population_IDF_Nettoyee_Enrichie.csv", index=False)

print("✅ Fichier 'Population_IDF_Nettoyee_Enrichie.csv' généré avec succès dans le dossier output_data.")
