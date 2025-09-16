"""
Script de debug pour comprendre pourquoi les bénéficiaires ne sont pas mappés correctement
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.file_handler import FileHandler
from core.beneficiary_resolver import BeneficiaryResolver
import pandas as pd
import logging

# Configuration des logs très détaillés
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - [%(levelname)s] - %(message)s'
)

print("\n" + "="*80)
print(" 🔍 DEBUG COMPLET DU MAPPING DES BÉNÉFICIAIRES")
print("="*80)

# Chemins des fichiers
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\BulkReport_130809.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Export_0131-FMC19-Beat.xlsx"

# Si les fichiers de test existent, les utiliser
test_bulk = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\BulkReport_Test.csv"
test_export = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\Export_Test.xlsx"

if os.path.exists(test_bulk):
    print(f"✓ Utilisation des fichiers de test")
    bulk_path = test_bulk
    export_path = test_export
else:
    print(f"✓ Utilisation des fichiers réels")

# 1. LECTURE DU BULKREPORT
print("\n" + "="*60)
print(" 📁 ÉTAPE 1: LECTURE DU BULKREPORT")
print("="*60)

handler = FileHandler()
bulk_df, metadata = handler.read_bulk_report(bulk_path)

print(f"\n✓ {len(bulk_df)} transactions lues")
print("\nColonnes disponibles dans BulkReport:")
for col in bulk_df.columns:
    print(f"  • {col}")

if len(bulk_df) > 0:
    print("\nPremière transaction:")
    row = bulk_df.iloc[0]
    for col in ['Credit Msisdn', 'TransactionID', 'Amount']:
        if col in bulk_df.columns:
            print(f"  {col}: {row[col]}")

# 2. LECTURE DU FICHIER EXPORT
print("\n" + "="*60)
print(" 📁 ÉTAPE 2: LECTURE DU FICHIER EXPORT")
print("="*60)

export_df = handler.read_export_file(export_path)

print(f"\n✓ {len(export_df)} bénéficiaires lus")
print("\nColonnes disponibles dans Export:")
for col in export_df.columns:
    print(f"  • {col}")

print("\n🔍 ANALYSE DES COLONNES DE NOMS:")
print("-"*40)

# Chercher les colonnes contenant des noms
name_columns = []
for col in export_df.columns:
    col_lower = col.lower()
    if any(word in col_lower for word in ['nom', 'name', 'prénom', 'prenom', 'beneficiaire']):
        name_columns.append(col)
        print(f"  ✓ Colonne trouvée: '{col}'")

if not name_columns and len(export_df.columns) > 0:
    print("  ⚠ Pas de colonne 'nom' trouvée, colonnes disponibles:")
    for col in export_df.columns:
        print(f"    - {col} (type: {export_df[col].dtype})")

# Afficher les premières lignes du DataFrame Export
print("\n📊 CONTENU DU FICHIER EXPORT (premières lignes):")
print("-"*40)
print(export_df.head(3).to_string())

# Afficher les bénéficiaires extraits
print("\n👥 BÉNÉFICIAIRES EXTRAITS:")
print("-"*40)

for i in range(min(5, len(export_df))):
    row = export_df.iloc[i]
    print(f"\nBénéficiaire {i+1}:")
    for col in export_df.columns:
        if pd.notna(row[col]):
            print(f"  {col}: {row[col]}")

# 3. TEST DU BENEFICIARY RESOLVER
print("\n" + "="*60)
print(" 🔧 ÉTAPE 3: TEST DU BENEFICIARY RESOLVER")
print("="*60)

resolver = BeneficiaryResolver()

print("\n📋 Détection de la stratégie de mapping...")
strategy = resolver._determine_strategy(bulk_df, export_df)
print(f"  Stratégie: {strategy['name']}")
print(f"  Raison: {strategy['reason']}")

print("\n🔄 Application du mapping...")
result_df = resolver.resolve_beneficiaries(bulk_df, export_df)

print(f"\n✓ Mapping terminé")
print(f"  • Colonnes du résultat: {list(result_df.columns)}")

# Vérifier la colonne Beneficiaire
if 'Beneficiaire' in result_df.columns:
    print("\n📊 RÉSULTAT DU MAPPING:")
    print("-"*40)
    
    for i in range(min(5, len(result_df))):
        row = result_df.iloc[i]
        beneficiaire = row.get('Beneficiaire', 'N/A')
        trans_id = row.get('TransactionID', 'N/A')
        phone = row.get('Credit Msisdn', 'N/A')
        
        print(f"\nTransaction {i+1}:")
        print(f"  ID: {trans_id}")
        print(f"  Téléphone: {phone}")
        print(f"  Bénéficiaire: {beneficiaire}")
        
        # Vérifier si c'est un placeholder
        if 'BENEFICIAIRE_' in str(beneficiaire):
            print(f"  ⚠️ ATTENTION: Placeholder détecté!")
else:
    print("\n❌ ERREUR: Colonne 'Beneficiaire' non trouvée dans le résultat!")

# 4. ANALYSE DES PROBLÈMES
print("\n" + "="*60)
print(" 🚨 ANALYSE DES PROBLÈMES")
print("="*60)

problems = []

# Vérifier les colonnes de noms dans Export
if not name_columns:
    problems.append("Pas de colonne de noms trouvée dans Export")

# Vérifier si les bénéficiaires sont des placeholders
if 'Beneficiaire' in result_df.columns:
    placeholders = result_df['Beneficiaire'].astype(str).str.contains('BENEFICIAIRE_').sum()
    if placeholders > 0:
        problems.append(f"{placeholders} bénéficiaires sont des placeholders")

# Vérifier le mapping par téléphone
phone_col_bulk = None
phone_col_export = None

for col in bulk_df.columns:
    if 'msisdn' in col.lower() or 'phone' in col.lower():
        phone_col_bulk = col
        break

for col in export_df.columns:
    if 'phone' in col.lower() or 'tel' in col.lower():
        phone_col_export = col
        break

if phone_col_bulk and not phone_col_export:
    problems.append(f"Colonne téléphone trouvée dans BulkReport ({phone_col_bulk}) mais pas dans Export")

if problems:
    print("\n⚠️ Problèmes détectés:")
    for p in problems:
        print(f"  • {p}")
else:
    print("\n✅ Aucun problème détecté")

# 5. RECOMMANDATIONS
print("\n" + "="*60)
print(" 💡 RECOMMANDATIONS")
print("="*60)

print("\n1. Structure attendue du fichier Export:")
print("   - Une colonne 'Nom et prénoms' ou 'Nom'")
print("   - Une colonne 'Téléphone' (optionnel)")

print("\n2. Vérifiez que le fichier Export contient bien:")
print("   - Les noms des bénéficiaires")
print("   - Le bon nombre de bénéficiaires")

print("\n3. Solutions possibles:")
print("   - S'assurer que la colonne des noms existe")
print("   - Vérifier l'encodage du fichier")
print("   - Vérifier que les cellules ne sont pas vides")

print("\n" + "="*80)
