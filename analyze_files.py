"""
Script d'analyse pour comprendre la structure des fichiers
"""
import pandas as pd
import os

base_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP"

print("="*60)
print("ANALYSE DÉTAILLÉE DES FICHIERS")
print("="*60)

# 1. Analyser Export Excel en détail
print("\n1. FICHIER EXPORT:")
print("-"*40)
try:
    file_path = os.path.join(base_path, "Export_0131-FMC19-Beat.xlsx")
    
    # Lire toutes les feuilles
    xl = pd.ExcelFile(file_path)
    print(f"Nombre de feuilles: {len(xl.sheet_names)}")
    print(f"Noms des feuilles: {xl.sheet_names}")
    
    # Analyser chaque feuille
    for sheet_name in xl.sheet_names:
        print(f"\n--- Feuille: {sheet_name} ---")
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        print(f"Dimensions: {df.shape}")
        
        # Chercher la colonne "Nom et prénoms"
        found = False
        for i in range(min(20, len(df))):
            row = df.iloc[i]
            for j, val in enumerate(row):
                if pd.notna(val) and "nom" in str(val).lower():
                    print(f"Trouvé 'nom' à la ligne {i}, colonne {j}: {val}")
                    found = True
                    
                    # Afficher les données autour
                    if i < len(df) - 5:
                        print(f"\nDonnées après cette ligne:")
                        for k in range(i+1, min(i+6, len(df))):
                            print(f"  Ligne {k}: {df.iloc[k, j]}")
                    break
            if found:
                break
                
        # Si pas trouvé, afficher un échantillon
        if not found:
            print("Pas de colonne 'nom' trouvée")
            print("\nÉchantillon des données non vides:")
            for i in range(min(30, len(df))):
                row_data = df.iloc[i].dropna().tolist()
                if row_data:
                    print(f"  Ligne {i}: {row_data[:5]}")  # Max 5 colonnes
                    
except Exception as e:
    print(f"Erreur: {e}")

# 2. Analyser le fichier frais
print("\n\n2. FICHIER FRAIS:")
print("-"*40)
try:
    file_path = os.path.join(base_path, "frais.xlsx")
    df = pd.read_excel(file_path, header=None)
    
    print(f"Dimensions: {df.shape}")
    print("\n10 premières lignes:")
    print(df.head(10))
    
    # Essayer avec header
    df2 = pd.read_excel(file_path)
    print("\nAvec header automatique:")
    print(f"Colonnes: {df2.columns.tolist()}")
    print(df2.head())
    
except Exception as e:
    print(f"Erreur: {e}")

print("\n" + "="*60)
