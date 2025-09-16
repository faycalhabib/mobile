"""
Script de debug pour identifier l'erreur tuple
"""
import sys
import traceback
from core.file_handler import FileHandler
from core.data_processor import DataProcessor
from core.report_generator import ReportGenerator

# Chemins des fichiers
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\BulkReport_130809.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Export_0131-FMC19-Beat.xlsx"
fees_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\frais.xlsx"

try:
    print("="*60)
    print("TEST DE GÉNÉRATION DE RAPPORT")
    print("="*60)
    
    # 1. Charger les fichiers
    handler = FileHandler()
    
    print("\n1. Chargement BulkReport...")
    bulk_df, metadata = handler.read_bulk_report(bulk_path)
    print(f"   ✓ {len(bulk_df)} transactions")
    print(f"   Colonnes: {bulk_df.columns.tolist()}")
    print(f"   Première ligne:")
    if len(bulk_df) > 0:
        for col in bulk_df.columns:
            val = bulk_df.iloc[0][col]
            print(f"     {col}: {val} (type: {type(val).__name__})")
    
    print("\n2. Chargement Export...")
    export_df = handler.read_export_file(export_path)
    print(f"   ✓ {len(export_df)} lignes")
    print(f"   Colonnes: {export_df.columns.tolist()}")
    
    print("\n3. Chargement Frais...")
    fees_df = handler.read_fees_file(fees_path)
    print(f"   ✓ {len(fees_df)} lignes")
    
    print("\n4. Traitement des données...")
    processor = DataProcessor()
    
    # Préparer les métadonnées
    metadata_report = {
        'date_paiement': '09/09/2025',
        'libelle': 'PAIEMENT LOCATION SALLE',
        'budget': 500000,
        'projet': 'UGP'
    }
    
    processed_df, errors = processor.process_transactions(
        bulk_df, export_df, fees_df, metadata_report
    )
    
    print(f"   ✓ {len(processed_df)} transactions traitées")
    print(f"   Colonnes finales: {processed_df.columns.tolist()}")
    
    # Afficher les erreurs
    if errors:
        print(f"\n   Avertissements:")
        for err in errors:
            print(f"     {err}")
    
    # Afficher le détail de la première transaction
    if len(processed_df) > 0:
        print(f"\n   Première transaction traitée:")
        for col in processed_df.columns:
            val = processed_df.iloc[0][col]
            print(f"     {col}: {val} (type: {type(val).__name__})")
    
    print("\n5. Génération du rapport...")
    generator = ReportGenerator()
    output_path = generator.generate_report(
        processed_df,
        metadata_report,
        "test_debug.xlsx"
    )
    
    print(f"   ✓ Rapport généré: {output_path}")
    
except Exception as e:
    print(f"\n❌ ERREUR: {e}")
    print(f"\nType d'erreur: {type(e).__name__}")
    print(f"\nTraceback complet:")
    traceback.print_exc()
