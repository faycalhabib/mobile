"""
Test direct d'écriture dans le template Excel
"""
import win32com.client as win32
import pythoncom
import os

template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
output_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\outputs\test_direct.xlsx"

pythoncom.CoInitialize()
try:
    # Copier le template
    import shutil
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.copy2(template_path, output_path)
    
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Open(os.path.abspath(output_path))
    sheet = wb.Worksheets('Rapport paiement')
    
    print("Test d'écriture directe...")
    
    # Métadonnées (lignes 6-9)
    sheet.Cells(6, 3).Value = "16/09/2025"  # Date de paiement
    sheet.Cells(7, 3).Value = "TEST DIRECT"  # Libellé
    sheet.Cells(8, 3).Value = "500 000"  # Budget
    
    # Transactions (ligne 12 et 13)
    # Ligne 12 - Transaction 1
    sheet.Cells(12, 2).Value = "09/09/2025"  # Date
    sheet.Cells(12, 3).Value = "CI9510O2KX"  # N° Transaction
    sheet.Cells(12, 4).Value = "PAIEMENT"  # Type
    sheet.Cells(12, 5).Value = "Success"  # Statut
    sheet.Cells(12, 6).Value = "491 741"  # Montant
    sheet.Cells(12, 7).Value = "8 261"  # Frais
    sheet.Cells(12, 8).Value = "UGP"  # De
    sheet.Cells(12, 9).Value = "23596771275"  # Vers
    sheet.Cells(12, 10).Value = "TINA GANG-IRANGA"  # Bénéficiaire
    
    # Ligne 13 - Transaction 2
    sheet.Cells(13, 2).Value = "09/09/2025"  # Date
    sheet.Cells(13, 3).Value = "CI95110BBF"  # N° Transaction
    sheet.Cells(13, 4).Value = "PAIEMENT"  # Type
    sheet.Cells(13, 5).Value = "Success"  # Statut
    sheet.Cells(13, 6).Value = "5 000"  # Montant
    sheet.Cells(13, 7).Value = "84"  # Frais
    sheet.Cells(13, 8).Value = "UGP"  # De
    sheet.Cells(13, 9).Value = "23596771275"  # Vers
    sheet.Cells(13, 10).Value = "TINA GANG-IRANGA"  # Bénéficiaire
    
    # Total ligne 15
    sheet.Cells(15, 5).Value = "TOTAL:"
    sheet.Cells(15, 6).Value = "496 741"
    sheet.Cells(15, 7).Value = "8 345"
    
    print("Données écrites:")
    print("  - Métadonnées (lignes 6-8)")
    print("  - Transaction 1 (ligne 12)")
    print("  - Transaction 2 (ligne 13)")
    print("  - Total (ligne 15)")
    
    wb.SaveAs(os.path.abspath(output_path), FileFormat=51)
    print(f"\n✓ Fichier sauvegardé: {output_path}")
    
    wb.Close(False)
    excel.Quit()
    
except Exception as e:
    print(f"Erreur: {e}")
    import traceback
    traceback.print_exc()
finally:
    pythoncom.CoUninitialize()
    
print("\nOuvrez le fichier test_direct.xlsx pour vérifier")
