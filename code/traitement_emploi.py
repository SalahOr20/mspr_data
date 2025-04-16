import pandas as pd


df = pd.read_excel("../Data/data_emploi.xlsx", sheet_name="Sheet1")

df = df.dropna(subset=["Trimestre", "Offres d'emploi"])
df = df.drop_duplicates()

df["Année"] = df["Trimestre"].astype(str).str.extract(r"(\d{4})")
df["Année"] = pd.to_numeric(df["Année"], errors="coerce")  # conversion sécurisée
df = df.dropna(subset=["Année"])                           # suppression des lignes non converties
df["Année"] = df["Année"].astype(int)


df = df[df["Année"].between(2010, 2024)]

df_annuel = df.groupby("Année", as_index=False).agg(
    Total_Offres=("Offres d'emploi", "sum")
)


base_2010 = df_annuel.loc[df_annuel["Année"] == 2010, "Total_Offres"].values[0]
df_annuel["Evolution_depuis_2010"] = df_annuel["Total_Offres"] - base_2010
df_annuel["Taux_Croissance_depuis_2010"] = (
    df_annuel["Evolution_depuis_2010"] / base_2010
) * 100
df_annuel["Moyenne_5_Dernieres"] = df_annuel["Total_Offres"].rolling(5).mean()

df_annuel.to_csv("../output_data/offre_emploi.csv", index=False)

print("✅ Fichier CSV généré : Offres_Emploi_Par_Annee.csv")