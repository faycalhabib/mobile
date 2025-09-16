"""
Test avec les nouveaux fichiers BulkReport et Export créés
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.file_handler import FileHandler
from core.data_processor import DataProcessor
import pandas as pd
import shutil

# Chemins des nouveaux fichiers de test
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\BulkReport_Test.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\test_data\Export_Test.xlsx"
template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
output_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\outputs\test_new_format.xlsx"

print("="*70)
print(" TEST AVEC LES NOUVEAUX FICHIERS")
print("="*70)

try:
    # 1. Charger les fichiers
    handler = FileHandler()
    
    print("\n📁 Chargement des fichiers...")
    bulk_df, metadata = handler.read_bulk_report(bulk_path)
    print(f"  ✓ BulkReport: {len(bulk_df)} transactions")
    
    if len(bulk_df) > 0:
        print("\n  Transactions chargées:")
        for i in range(len(bulk_df)):
            print(f"    {i+1}. {bulk_df.iloc[i]['Credit Msisdn']} - {bulk_df.iloc[i]['Amount']} FCFA - ID: {bulk_df.iloc[i]['TransactionID']}")
    
    export_df = handler.read_export_file(export_path)
    print(f"\n  ✓ Export: {len(export_df)} bénéficiaires")
    
    if len(export_df) > 0:
        print("\n  Bénéficiaires chargés:")
        for i in range(len(export_df)):
            nom = export_df.iloc[i].get('Nom', export_df.iloc[i].get('Nom et prénoms', 'N/A'))
            tel = export_df.iloc[i].get('Telephone', export_df.iloc[i].get('Téléphone', 'N/A'))
            print(f"    {i+1}. {nom} - {tel}")
    
    # 2. Traiter les données
    print("\n🔄 Traitement des données...")
    processor = DataProcessor()
    processor.use_smart_processing = True
    
    metadata_report = {
        'date_paiement': '16/09/2025',
        'libelle': 'PAIEMENT SALAIRE',
        'budget': 2500000,
        'projet': 'UGP'
    }
    
    processed_df, errors = processor.process_transactions(
        bulk_df, export_df, pd.DataFrame(), metadata_report
    )
    
    print(f"\n  ✓ Transactions traitées: {len(processed_df)}")
    
    # Afficher le mapping
    if len(processed_df) > 0:
        print("\n📊 RÉSULTAT DU MAPPING:")
        print("-"*60)
        for i in range(len(processed_df)):
            row = processed_df.iloc[i]
            print(f"\n  Transaction {i+1}:")
            print(f"    • ID: {row['TransactionID']}")
            print(f"    • Montant: {row['Amount']:,.0f} FCFA")
            print(f"    • Vers: {row['Vers']}")
            print(f"    • Bénéficiaire: {row['Beneficiaire']}")
            print(f"    • Frais: {row['Frais']:,.0f} FCFA")
    
    # 3. Générer le rapport Excel
    print("\n📝 Génération du rapport Excel...")
    
    # Copier le template
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.copy2(template_path, output_path)
    
    # Écrire les données avec win32com
    import win32com.client as win32
    import pythoncom
    
    pythoncom.CoInitialize()
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Open(os.path.abspath(output_path))
    sheet = wb.Worksheets('Rapport paiement')
    
    # Métadonnées
    sheet.Cells(6, 3).Value = metadata_report['date_paiement']
    sheet.Cells(7, 3).Value = metadata_report['libelle']
    sheet.Cells(8, 3).Value = str(metadata_report['budget'])
    
    # Écrire les transactions
    start_row = 12
    for idx, row in processed_df.iterrows():
        excel_row = start_row + idx
        
        sheet.Cells(excel_row, 2).Value = str(row.get('Date', ''))
        sheet.Cells(excel_row, 3).Value = str(row.get('TransactionID', ''))
        sheet.Cells(excel_row, 4).Value = 'PAIEMENT'
        sheet.Cells(excel_row, 5).Value = str(row.get('Status', 'Success'))
        sheet.Cells(excel_row, 6).Value = str(int(row.get('Amount', 0)))
        sheet.Cells(excel_row, 7).Value = str(int(row.get('Frais', 0)))
        sheet.Cells(excel_row, 8).Value = 'UGP'
        sheet.Cells(excel_row, 9).Value = str(row.get('Vers', ''))
        sheet.Cells(excel_row, 10).Value = str(row.get('Beneficiaire', ''))
        
        print(f"    Ligne {excel_row}: {row.get('TransactionID')} → {row.get('Beneficiaire')}")
    
    # Total
    if len(processed_df) > 0:
        total_row = start_row + len(processed_df) + 1
        sheet.Cells(total_row, 5).Value = "TOTAL:"
        sheet.Cells(total_row, 6).Value = str(int(processed_df['Amount'].sum()))
        sheet.Cells(total_row, 7).Value = str(int(processed_df['Frais'].sum()))
    
    wb.Save()
    wb.Close(False)
    excel.Quit()
    pythoncom.CoUninitialize()
    
    print(f"\n✅ Rapport généré: {output_path}")
    print("\nOuvrez le fichier pour vérifier le résultat!")
    
except Exception as e:
    print(f"\n❌ Erreur: {e}")
    import traceback
    traceback.print_exc()
