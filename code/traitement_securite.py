import pandas as pd

# 1. Charger le fichier
df = pd.read_excel("../Data/data_securitee.xlsx")

# 2. Nettoyer les noms de colonnes
df.columns = [col.strip().replace("'", "").replace("**", "").lower() for col in df.columns]

# 3. Renommer les colonnes (dont 'année' → 'annee')
df.rename(columns={
    "année": "annee",
    "type dinfraction": "type_infraction",
    "nombre dinfractions": "nombre_infractions",
    "taux pour 1000 hab.": "taux_pour_1000_hab",
    "région": "region"
}, inplace=True)

# 4. Forcer les types
df["annee"] = df["annee"].astype(int)
df["nombre_infractions"] = pd.to_numeric(df["nombre_infractions"], errors="coerce")
df["taux_pour_1000_hab"] = pd.to_numeric(df["taux_pour_1000_hab"], errors="coerce")
df["region"] = df["region"].astype(str)
df["type_infraction"] = df["type_infraction"].astype(str)

# 5. Filtrer Île-de-France uniquement
df_idf = df[df["region"].str.lower().str.contains("île-de-france")]

# 6. Export final
df_idf.to_csv("../output_data/securite_idf.csv", index=False)

print("✅ Fichier 'securite_idf_2016_2023.csv' exporté avec succès dans output_data.")
