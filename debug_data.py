"""
Script de debug pour vérifier que les données sont bien transmises
"""
from core.file_handler import FileHandler
from core.data_processor import DataProcessor
import pandas as pd

# Chemins
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\BulkReport_130809.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Export_0131-FMC19-Beat.xlsx"
fees_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\frais.xlsx"

try:
    print("="*60)
    print("DEBUG: VÉRIFICATION DES DONNÉES")
    print("="*60)
    
    handler = FileHandler()
    
    # 1. Charger BulkReport
    print("\n1. BulkReport:")
    bulk_df, metadata = handler.read_bulk_report(bulk_path)
    print(f"   Nombre de lignes après filtrage: {len(bulk_df)}")
    if len(bulk_df) > 0:
        print(f"   Colonnes: {bulk_df.columns.tolist()}")
        print(f"   Première transaction:")
        print(bulk_df.iloc[0])
    
    # 2. Charger Export
    print("\n2. Export:")
    export_df = handler.read_export_file(export_path)
    print(f"   Nombre de bénéficiaires: {len(export_df)}")
    
    # 3. Traiter les données
    print("\n3. Traitement:")
    processor = DataProcessor()
    metadata_report = {
        'date_paiement': '16/09/2025',
        'libelle': 'TEST',
        'budget': 500000,
        'projet': 'UGP'
    }
    
    processed_df, errors = processor.process_transactions(
        bulk_df, export_df, 
        handler.read_fees_file(fees_path),
        metadata_report
    )
    
    print(f"   Transactions traitées: {len(processed_df)}")
    print(f"   Colonnes finales: {processed_df.columns.tolist()}")
    
    if len(processed_df) > 0:
        print("\n   Première transaction traitée:")
        for col in processed_df.columns:
            print(f"     {col}: {processed_df.iloc[0][col]}")
    
    # Vérifier les colonnes essentielles
    print("\n4. Vérification des colonnes essentielles:")
    essential_cols = ['Date', 'TransactionID', 'Status', 'Amount', 'Frais', 'Vers', 'Beneficiaire']
    for col in essential_cols:
        if col in processed_df.columns:
            print(f"   ✓ {col} présent")
        else:
            print(f"   ✗ {col} MANQUANT!")
    
except Exception as e:
    print(f"\n❌ Erreur: {e}")
    import traceback
    traceback.print_exc()
