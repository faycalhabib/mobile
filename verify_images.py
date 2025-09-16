"""
Script pour vérifier la présence d'images dans le template
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
    
    for i in range(1, wb.Worksheets.Count + 1):
        sheet = wb.Worksheets(i)
        print(f"\nFeuille '{sheet.Name}':")
        
        # Compter les formes (shapes) qui incluent les images
        shape_count = sheet.Shapes.Count
        print(f"  Nombre de formes/images: {shape_count}")
        
        if shape_count > 0:
            for j in range(1, min(shape_count + 1, 6)):  # Afficher max 5
                shape = sheet.Shapes(j)
                print(f"    - {shape.Name} (Type: {shape.Type})")
    
    wb.Close(False)
    excel.Quit()
    
except Exception as e:
    print(f"Erreur: {e}")
finally:
    pythoncom.CoUninitialize()
