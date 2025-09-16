"""
Script de débogage pour identifier la structure exacte du template
"""
import win32com.client as win32
import pythoncom
import os

template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"

pythoncom.CoInitialize()
try:
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    
    wb = excel.Workbooks.Open(os.path.abspath(template_path))
    sheet = wb.Worksheets('Rapport paiement')
    
    print("="*60)
    print("ANALYSE DU TEMPLATE - Rapport paiement")
    print("="*60)
    
    # Chercher l'en-tête du tableau
    print("\nRecherche de l'en-tête 'Date':")
    for row in range(1, 20):
        for col in range(1, 12):
            cell_value = sheet.Cells(row, col).Value
            if cell_value:
                cell_str = str(cell_value).strip()
                if 'Date' in cell_str or 'date' in cell_str.lower():
                    print(f"  Trouvé 'Date' en ligne {row}, colonne {col}: '{cell_str}'")
                    
                    # Afficher toute la ligne d'en-tête
                    print(f"\n  En-têtes ligne {row}:")
                    for c in range(1, 12):
                        val = sheet.Cells(row, c).Value
                        if val:
                            print(f"    Col {c}: {val}")
                    
                    # Vérifier les lignes suivantes (où les données devraient être)
                    print(f"\n  Lignes suivantes (où mettre les données):")
                    for r in range(row+1, min(row+5, 20)):
                        print(f"    Ligne {r}:")
                        has_data = False
                        for c in range(1, 10):
                            val = sheet.Cells(r, c).Value
                            if val:
                                print(f"      Col {c}: {val}")
                                has_data = True
                        if not has_data:
                            print(f"      [VIDE - OK pour remplir les données]")
                    break
        else:
            continue
        break
    
    # Chercher les métadonnées
    print("\n\nMétadonnées trouvées:")
    for row in range(1, 10):
        for col in range(1, 5):
            cell_value = sheet.Cells(row, col).Value
            if cell_value:
                cell_str = str(cell_value)
                if any(x in cell_str.lower() for x in ['date de paiement', 'libellé', 'budget', 'projet']):
                    next_val = sheet.Cells(row, col+1).Value
                    print(f"  Ligne {row}, Col {col}: {cell_str} -> Col {col+1}: {next_val}")
    
    wb.Close(False)
    excel.Quit()
    
except Exception as e:
    print(f"Erreur: {e}")
    import traceback
    traceback.print_exc()
finally:
    pythoncom.CoUninitialize()
