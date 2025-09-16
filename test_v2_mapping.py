"""
Test du nouveau système de mapping V2
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.file_handler import FileHandler
from core.beneficiary_resolver_v2 import BeneficiaryResolverV2
import pandas as pd
import logging

# Logs détaillés
logging.basicConfig(
    level=logging.INFO,
    format='%(message)s'
)

print("\n" + "="*80)
print(" TEST DU SYSTÈME DE MAPPING V2")
print("="*80)

# Fichiers de test
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\BulkReport_Test.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\Export_Test.xlsx"

# Charger les fichiers
handler = FileHandler()
bulk_df, _ = handler.read_bulk_report(bulk_path)
export_df = handler.read_export_file(export_path)

print(f"\n✓ BulkReport: {len(bulk_df)} transactions")
print(f"✓ Export: {len(export_df)} bénéficiaires")

# Tester le nouveau resolver
print("\n" + "="*60)
print(" TEST DU BENEFICIARY RESOLVER V2")
print("="*60)

resolver = BeneficiaryResolverV2()
result = resolver.resolve_beneficiaries(bulk_df, export_df)

# Vérifier le résultat
print("\n" + "="*60)
print(" VÉRIFICATION DU RÉSULTAT")
print("="*60)

if 'Beneficiaire' in result.columns:
    print("\n✅ Colonne 'Beneficiaire' créée")
    
    # Analyser les bénéficiaires
    placeholders = 0
    real_names = 0
    
    print("\nBénéficiaires mappés:")
    for i in range(len(result)):
        benef = result.iloc[i]['Beneficiaire']
        trans_id = result.iloc[i].get('TransactionID', 'N/A')
        
        if 'BENEFICIAIRE_' in str(benef):
            placeholders += 1
            status = "❌ PLACEHOLDER"
        else:
            real_names += 1
            status = "✅"
        
        print(f"  {i+1}. {trans_id} → {benef} {status}")
    
    print(f"\nRésumé:")
    print(f"  • Vrais noms: {real_names}/{len(result)}")
    print(f"  • Placeholders: {placeholders}/{len(result)}")
    
    if placeholders > 0:
        print("\n⚠️ Des placeholders sont encore présents!")
        print("Vérifiez que le fichier Export contient bien une colonne 'Nom et prénoms'")
    else:
        print("\n✅ SUCCÈS! Tous les bénéficiaires ont des vrais noms!")
else:
    print("\n❌ ERREUR: Colonne 'Beneficiaire' non créée")

print("\n" + "="*80)
