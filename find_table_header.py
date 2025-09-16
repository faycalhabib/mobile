"""
Chercher le VRAI en-tête du tableau des transactions
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
    print("RECHERCHE DU TABLEAU DES TRANSACTIONS")
    print("="*60)
    
    # Chercher spécifiquement les colonnes du tableau
    # En regardant l'image, on voit: Date | N° Transaction | Type | Statut | Montant | Frais ONG | De | Vers | Bénéficiaire
    
    print("\nRecherche des en-têtes du tableau (N° Transaction, Type, Statut, etc.):")
    
    for row in range(1, 30):  # Chercher plus loin
        found_headers = []
        for col in range(1, 15):
            cell_value = sheet.Cells(row, col).Value
            if cell_value:
                cell_str = str(cell_value).strip()
                # Chercher les marqueurs du tableau
                if any(marker in cell_str for marker in ['N° Transaction', 'Type', 'Statut', 'Montant', 'Frais ONG', 'Bénéficiaire']):
                    found_headers.append((col, cell_str))
        
        if len(found_headers) >= 3:  # Si on trouve au moins 3 en-têtes, c'est notre ligne
            print(f"\n✓ TROUVÉ! Ligne {row} contient les en-têtes du tableau:")
            print(f"  En-têtes trouvés: {found_headers}")
            
            # Afficher TOUS les en-têtes de cette ligne
            print(f"\n  Tous les en-têtes de la ligne {row}:")
            for c in range(1, 15):
                val = sheet.Cells(row, c).Value
                if val:
                    print(f"    Colonne {c}: '{val}'")
            
            # Vérifier les lignes suivantes pour le remplissage
            print(f"\n  Lignes suivantes (pour remplir les données):")
            for r in range(row+1, min(row+6, 30)):
                print(f"    Ligne {r}:", end="")
                has_content = False
                for c in range(1, 10):
                    val = sheet.Cells(r, c).Value
                    if val:
                        has_content = True
                        break
                if has_content:
                    print(" [Contient des données]")
                else:
                    print(" [VIDE - Prêt pour les données]")
            break
    else:
        print("\n⚠ Tableau non trouvé avec les marqueurs attendus")
        print("\nAnalyse ligne par ligne (lignes 10-20):")
        for row in range(10, 21):
            print(f"\nLigne {row}:")
            for col in range(1, 10):
                val = sheet.Cells(row, col).Value
                if val:
                    print(f"  Col {col}: {val}")
    
    wb.Close(False)
    excel.Quit()
    
except Exception as e:
    print(f"Erreur: {e}")
finally:
    pythoncom.CoUninitialize()
