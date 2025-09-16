"""
Test du nouveau système d'écriture intelligente ExcelSmartWriter
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.file_handler import FileHandler
from core.data_processor import DataProcessor
from core.excel_smart_writer import ExcelSmartWriter
import pandas as pd
import shutil
import logging

# Configuration des logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

# Chemins des fichiers
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\BulkReport_Test.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\Export_Test.xlsx"
template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
output_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\outputs\test_smart_writer.xlsx"

print("\n" + "="*70)
print(" TEST DU SYSTÈME D'ÉCRITURE INTELLIGENTE")
print("="*70)

try:
    # 1. Charger les données
    handler = FileHandler()
    
    print("\n📁 Chargement des données...")
    bulk_df, metadata = handler.read_bulk_report(bulk_path)
    print(f"  ✓ BulkReport: {len(bulk_df)} transactions")
    
    export_df = handler.read_export_file(export_path)
    print(f"  ✓ Export: {len(export_df)} bénéficiaires")
    
    # 2. Traiter les données
    print("\n🔄 Traitement des données...")
    processor = DataProcessor()
    processor.use_smart_processing = True
    
    metadata_report = {
        'date_paiement': '16/09/2025',
        'libelle': 'PAIEMENT SALAIRE OCTOBRE',
        'budget': 2500000,
        'projet': 'UGP'
    }
    
    processed_df, errors = processor.process_transactions(
        bulk_df, export_df, pd.DataFrame(), metadata_report
    )
    
    print(f"  ✓ {len(processed_df)} transactions traitées")
    
    # 3. Copier le template
    print("\n📋 Préparation du template...")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.copy2(template_path, output_path)
    print(f"  ✓ Template copié")
    
    # 4. Utiliser ExcelSmartWriter
    print("\n✍️ Écriture intelligente dans Excel...")
    writer = ExcelSmartWriter()
    success = writer.write_report(output_path, processed_df, metadata_report)
    
    if success:
        print(f"\n✅ SUCCÈS! Rapport généré: {output_path}")
        print("\nCaractéristiques du rapport:")
        print(f"  • {len(processed_df)} transactions")
        print(f"  • Insertion automatique de lignes si > 2 transactions")
        print(f"  • Section Récapitulatif préservée")
        print(f"  • Format respecté")
    else:
        print("\n❌ Échec de la génération")
        
except Exception as e:
    print(f"\n❌ Erreur: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "="*70)
