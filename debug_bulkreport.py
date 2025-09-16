"""
Debug de la lecture du BulkReport pour identifier où les lignes sont perdues
"""
import pandas as pd
import chardet

bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\BulkReport_130809.csv"

print("="*70)
print(" DEBUG LECTURE BULKREPORT")
print("="*70)

# 1. Détecter l'encodage
with open(bulk_path, 'rb') as f:
    result = chardet.detect(f.read())
    encoding = result['encoding'] or 'utf-8'
print(f"\n1. Encodage détecté: {encoding}")

# 2. Chercher la ligne de départ
markers = ['Record No', 'Validation Result', 'Credit Msisdn', 'Transaction Timestamp']
data_start_line = 12  # Valeur par défaut

with open(bulk_path, 'r', encoding=encoding) as f:
    lines = f.readlines()
    for i, line in enumerate(lines):
        if any(marker in line for marker in markers):
            data_start_line = i
            print(f"\n2. Headers trouvés à la ligne {i+1} (index {i}):")
            print(f"   {line[:100]}...")
            break

# 3. Lire le CSV avec différentes stratégies
print(f"\n3. Lecture avec skiprows={data_start_line}:")
try:
    df = pd.read_csv(bulk_path, 
                     skiprows=data_start_line,
                     encoding=encoding,
                     on_bad_lines='skip')
    
    print(f"   • Lignes lues: {len(df)}")
    print(f"   • Colonnes: {list(df.columns[:5])}...")
    
    # Montrer les premières lignes
    if len(df) > 0:
        print("\n   Première ligne:")
        print(df.iloc[0])
    
    # 4. Vérifier les filtres appliqués
    print("\n4. Application des filtres:")
    
    # Nettoyer les colonnes
    df.columns = [col.strip() for col in df.columns]
    print(f"   • Après nettoyage colonnes: {len(df)} lignes")
    
    # Filtrer les lignes vides
    df_before = len(df)
    df = df.dropna(how='all')
    print(f"   • Après dropna(how='all'): {len(df)} lignes (perdu {df_before - len(df)})")
    
    # Vérifier la colonne Credit Msisdn
    if 'Credit Msisdn' in df.columns:
        na_count = df['Credit Msisdn'].isna().sum()
        print(f"   • NaN dans 'Credit Msisdn': {na_count}")
        
        df_before = len(df)
        df = df[df['Credit Msisdn'].notna()]
        print(f"   • Après filtre Credit Msisdn notna: {len(df)} lignes (perdu {df_before - len(df)})")
    else:
        print(f"   ⚠ Colonne 'Credit Msisdn' non trouvée!")
        print(f"   Colonnes disponibles: {list(df.columns)}")
    
    # Vérifier Amount
    if 'Amount' in df.columns:
        df_before = len(df)
        df = df.dropna(subset=['Amount'])
        print(f"   • Après dropna Amount: {len(df)} lignes (perdu {df_before - len(df)})")
        
        df_before = len(df)
        df = df[df['Amount'] > 0]
        print(f"   • Après Amount > 0: {len(df)} lignes (perdu {df_before - len(df)})")
    
    print(f"\n5. Résultat final: {len(df)} lignes")
    
    if len(df) > 0:
        print("\n6. Données finales:")
        for i in range(min(2, len(df))):
            print(f"\n   Transaction {i+1}:")
            for col in ['Credit Msisdn', 'Amount', 'TransactionID']:
                if col in df.columns:
                    print(f"     {col}: {df.iloc[i][col]}")
    
except Exception as e:
    print(f"   ❌ Erreur: {e}")
    import traceback
    traceback.print_exc()

# 7. Lecture alternative sans filtres
print("\n7. Lecture SANS AUCUN FILTRE:")
try:
    df_raw = pd.read_csv(bulk_path, encoding=encoding)
    print(f"   • Total lignes (sans skip): {len(df_raw)}")
    
    # Chercher manuellement les transactions
    for i, row in df_raw.iterrows():
        # Chercher des lignes qui ressemblent à des transactions
        row_str = str(row.values)
        if '23596771275' in row_str or 'CI9510O2KX' in row_str:
            print(f"\n   Trouvé transaction à ligne {i}:")
            print(f"   {row.values[:5]}...")
            
except Exception as e:
    print(f"   Erreur: {e}")
