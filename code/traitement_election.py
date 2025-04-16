import pandas as pd
import unicodedata
import os

# 🔧 Nettoyage des noms de colonnes
def clean_columns(columns):
    return [
        unicodedata.normalize('NFKD', col)
        .encode('ascii', 'ignore').decode('utf-8')
        .strip().lower().replace(" ", "_").replace("%", "pct")
        for col in columns
    ]

# 📥 Traitement d'une feuille d’élection
def lire_feuille_election(path, annee, tour):
    try:
        sheet = f"data_election_tour{tour}"
        print(f"📄 Lecture de {sheet} dans {os.path.basename(path)}")
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception as e:
        print(f"⚠️ Feuille {sheet} introuvable dans {path} → {e}")
        return pd.DataFrame()

    df.columns = clean_columns(df.columns)
    df = df.dropna(axis=1, how="all")
    df["annee"] = annee
    df["tour"] = tour

    # Calculs d’indicateurs utiles
    if "votants" in df.columns and "inscrits" in df.columns:
        df["taux_participation"] = df["votants"] / df["inscrits"]
    if "abstentions" in df.columns and "inscrits" in df.columns:
        df["taux_abstention"] = df["abstentions"] / df["inscrits"]
    if "exprimes" in df.columns and "inscrits" in df.columns:
        df["taux_exprimes"] = df["exprimes"] / df["inscrits"]

    # Moyenne et total voix candidats
    voix_cols = [col for col in df.columns if col.startswith("voix")]
    if voix_cols:
        df["total_voix_candidats"] = df[voix_cols].sum(axis=1, numeric_only=True)
        df["moyenne_voix_candidats"] = df[voix_cols].mean(axis=1, numeric_only=True)
        if "voix" in df.columns and "voix.1" in df.columns:
            df["ecart_voix_candidat_1_2"] = df["voix"] - df["voix.1"]

    return df

# 📁 Liste des fichiers d’élections
election_files = {
    2002: "C:/Users/Salah/Desktop/Cours/MSPR/Data/data_election_2002.xlsx",
    2007: "C:/Users/Salah/Desktop/Cours/MSPR/Data/data_election_2007.xlsx",
    2012: "C:/Users/Salah/Desktop/Cours/MSPR/Data/data_election_2012.xlsx",
    2017: "C:/Users/Salah/Desktop/Cours/MSPR/Data/data_election_2017.xlsx",
    2022: "C:/Users/Salah/Desktop/Cours/MSPR/Data/data_election_2022.xlsx"
}

# 🔄 Traitement de tous les fichiers
df_all = []
for annee, path in election_files.items():
    for tour in [1, 2]:
        if os.path.exists(path):
            df = lire_feuille_election(path, annee, tour)
            if not df.empty:
                df_all.append(df)
        else:
            print(f"❌ Fichier manquant : {path}")

# 🔀 Fusion
df_final = pd.concat(df_all, ignore_index=True)

# ✅ FILTRAGE : Départements d’Île-de-France uniquement
idf_departements = ['75', '77', '78', '91', '92', '93', '94', '95']
df_final["code_departement"] = df_final["code_departement"].astype(str).str.zfill(2)
df_idf = df_final[df_final["code_departement"].isin(idf_departements)]

# 💾 Export
df_idf.to_csv("election_global_idf_filtre.csv", index=False)
print("✅ Fichier exporté : election_global_idf_filtre.csv")
