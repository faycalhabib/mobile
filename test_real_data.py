"""
Test avec les vraies données du BulkReport_130809.csv
"""
import win32com.client as win32
import pythoncom
import os
import pandas as pd
from core.file_handler import FileHandler
from core.data_processor import DataProcessor

template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
output_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\outputs\test_real_data.xlsx"
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\BulkReport_130809.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Export_0131-FMC19-Beat.xlsx"

try:
    # Charger les données
    handler = FileHandler()
    bulk_df, metadata = handler.read_bulk_report(bulk_path)
    export_df = handler.read_export_file(export_path)
    
    print(f"Transactions chargées: {len(bulk_df)}")
    print(f"Bénéficiaires chargés: {len(export_df)}")
    
    # Traiter les données
    processor = DataProcessor()
    processed_df, errors = processor.process_transactions(
        bulk_df, export_df, pd.DataFrame(), 
        {'date_paiement': '16/09/2025', 'libelle': 'TEST', 'budget': 500000, 'projet': 'UGP'}
    )
    
    print(f"\nTransactions traitées: {len(processed_df)}")
    
    # Copier template et ouvrir avec Excel
    import shutil
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.copy2(template_path, output_path)
    
    pythoncom.CoInitialize()
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Open(os.path.abspath(output_path))
    sheet = wb.Worksheets('Rapport paiement')
    
    print("\nÉcriture dans Excel...")
    
    # Métadonnées
    sheet.Cells(6, 3).Value = "16/09/2025"
    sheet.Cells(7, 3).Value = "PAIEMENT LOCATION"
    sheet.Cells(8, 3).Value = "500 000"
    
    # Écrire les transactions (ligne 12+)
    start_row = 12
    for idx, record in processed_df.iterrows():
        row = start_row + idx
        print(f"\nLigne {row}:")
        print(f"  Date: {record.get('Date', '')}")
        print(f"  TransactionID: {record.get('TransactionID', '')}")
        print(f"  Amount: {record.get('Amount', 0)}")
        print(f"  Vers: {record.get('Vers', '')}")
        print(f"  Beneficiaire: {record.get('Beneficiaire', '')}")
        
        # Écrire dans Excel
        sheet.Cells(row, 2).Value = str(record.get('Date', ''))
        sheet.Cells(row, 3).Value = str(record.get('TransactionID', ''))
        sheet.Cells(row, 4).Value = "PAIEMENT"
        sheet.Cells(row, 5).Value = str(record.get('Status', 'Success'))
        sheet.Cells(row, 6).Value = str(int(record.get('Amount', 0)))
        sheet.Cells(row, 7).Value = str(int(record.get('Frais', 0)))
        sheet.Cells(row, 8).Value = "UGP"
        sheet.Cells(row, 9).Value = str(record.get('Vers', ''))
        sheet.Cells(row, 10).Value = str(record.get('Beneficiaire', ''))
    
    # Total
    if len(processed_df) > 0:
        total_row = start_row + len(processed_df) + 1
        sheet.Cells(total_row, 5).Value = "TOTAL:"
        sheet.Cells(total_row, 6).Value = str(int(processed_df['Amount'].sum()))
        sheet.Cells(total_row, 7).Value = str(int(processed_df['Frais'].sum()))
        print(f"\nTotal ligne {total_row}: {processed_df['Amount'].sum()}")
    
    wb.SaveAs(os.path.abspath(output_path), FileFormat=51)
    wb.Close(False)
    excel.Quit()
    
    print(f"\n✓ Fichier créé: {output_path}")
    
except Exception as e:
    print(f"Erreur: {e}")
    import traceback
    traceback.print_exc()
finally:
    pythoncom.CoUninitialize()
