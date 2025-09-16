"""
Test complet du système pour vérifier que toutes les données sont bien extraites
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.file_handler import FileHandler
from core.data_processor import DataProcessor
import pandas as pd
import logging

# Configuration des logs
logging.basicConfig(
    level=logging.INFO,
    format='%(message)s'
)

print("\n" + "="*70)
print(" TEST COMPLET DU SYSTÈME")
print("="*70)

# Chemins des fichiers
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\BulkReport_Test.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\Export_Test.xlsx"

# 1. Charger et vérifier BulkReport
print("\n📁 LECTURE DU BULKREPORT")
print("-"*60)

handler = FileHandler()
bulk_df, metadata = handler.read_bulk_report(bulk_path)

print(f"Nombre de transactions: {len(bulk_df)}")
print("\nColonnes disponibles:")
for col in bulk_df.columns:
    print(f"  • {col}")

if len(bulk_df) > 0:
    print("\n🔍 DONNÉES EXTRAITES (première transaction):")
    row = bulk_df.iloc[0]
    
    # Vérifier les colonnes importantes
    important_cols = ['Transaction Timestamp', 'Finished Timestamp', 'Credit Msisdn', 
                     'TransactionID', 'Amount', 'Status']
    
    for col in important_cols:
        if col in bulk_df.columns:
            print(f"  {col}: {row[col]}")
        else:
            print(f"  {col}: ⚠️ COLONNE MANQUANTE")

# 2. Charger et vérifier Export
print("\n📁 LECTURE DU FICHIER EXPORT")
print("-"*60)

export_df = handler.read_export_file(export_path)
print(f"Nombre de bénéficiaires: {len(export_df)}")
print("\nColonnes disponibles:")
for col in export_df.columns:
    print(f"  • {col}")

if len(export_df) > 0:
    print("\n🔍 BÉNÉFICIAIRES EXTRAITS:")
    for i in range(min(3, len(export_df))):
        row = export_df.iloc[i]
        name = row.get('Nom', row.get('Nom et prénoms', 'N/A'))
        phone = row.get('Telephone', row.get('Téléphone', 'N/A'))
        print(f"  {i+1}. {name} - Tel: {phone}")

# 3. Traiter avec DataProcessor
print("\n🔄 TRAITEMENT DES DONNÉES")
print("-"*60)

processor = DataProcessor()
processor.use_smart_processing = True

metadata_report = {
    'date_paiement': '17/09/2025',
    'libelle': 'TEST COMPLET',
    'budget': 2500000,
    'projet': 'UGP'
}

processed_df, errors = processor.process_transactions(
    bulk_df, export_df, pd.DataFrame(), metadata_report
)

print(f"Transactions traitées: {len(processed_df)}")

# 4. Vérifier les colonnes finales
print("\n✅ VÉRIFICATION DES COLONNES FINALES")
print("-"*60)

required_cols = ['Date', 'TransactionID', 'Type', 'Status', 'Amount', 
                'Frais', 'De', 'Vers', 'Beneficiaire']

for col in required_cols:
    if col in processed_df.columns:
        # Vérifier si la colonne a des valeurs
        non_empty = processed_df[col].notna().sum()
        if non_empty > 0:
            sample = processed_df[col].iloc[0]
            print(f"  ✓ {col}: {sample}")
        else:
            print(f"  ⚠️ {col}: COLONNE VIDE!")
    else:
        print(f"  ❌ {col}: COLONNE MANQUANTE!")

# 5. Afficher le mapping final
print("\n📊 MAPPING FINAL DES TRANSACTIONS")
print("-"*60)

for i in range(len(processed_df)):
    row = processed_df.iloc[i]
    print(f"\nTransaction {i+1}:")
    print(f"  • Date: {row.get('Date', 'N/A')}")
    print(f"  • ID: {row.get('TransactionID', 'N/A')}")
    print(f"  • Montant: {row.get('Amount', 0):,.0f} FCFA")
    print(f"  • Vers: {row.get('Vers', 'N/A')}")
    print(f"  • Bénéficiaire: {row.get('Beneficiaire', 'N/A')}")
    
    # Vérifier les problèmes
    if pd.isna(row.get('Date')) or row.get('Date') == '':
        print(f"    ⚠️ Date manquante!")
    if pd.isna(row.get('Vers')) or row.get('Vers') == '':
        print(f"    ⚠️ Numéro 'Vers' manquant!")
    if 'BENEFICIAIRE_' in str(row.get('Beneficiaire', '')):
        print(f"    ⚠️ Bénéficiaire non mappé correctement!")

# 6. Afficher les erreurs
if errors:
    print("\n⚠️ ERREURS/AVERTISSEMENTS:")
    for error in errors:
        print(f"  • {error}")

print("\n" + "="*70)
