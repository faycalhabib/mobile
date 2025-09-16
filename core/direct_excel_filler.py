"""
Solution GARANTIE : Ouvre directement le template original dans Excel
et le sauvegarde sous un nouveau nom après remplissage
"""
import os
import win32com.client as win32
import pythoncom
import pandas as pd
import logging
from datetime import datetime

logger = logging.getLogger(__name__)


class DirectExcelFiller:
    """Utilise Excel directement sans copier le fichier - Préserve TOUT"""
    
    def __init__(self):
        self.template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
        
    def fill_template(self, processed_df: pd.DataFrame, metadata: dict, output_path: str) -> str:
        """
        Ouvre le template ORIGINAL, remplit les données, sauvegarde sous nouveau nom
        GARANTIT 100% la préservation de TOUT
        """
        pythoncom.CoInitialize()
        excel = None
        workbook = None
        
        try:
            # Créer le dossier de sortie
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Ouvrir Excel
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # IMPORTANT: Ouvrir le template ORIGINAL (sans le copier)
            workbook = excel.Workbooks.Open(os.path.abspath(self.template_path))
            logger.info("Template original ouvert dans Excel")
            
            # Obtenir la feuille
            try:
                sheet = workbook.Worksheets('Rapport paiement')
            except:
                sheet = workbook.Worksheets(1)
            
            # Remplir les données
            self._fill_data(sheet, processed_df, metadata)
            
            # SAUVEGARDER SOUS UN NOUVEAU NOM (préserve tout)
            workbook.SaveAs(os.path.abspath(output_path))
            logger.info(f"Sauvegardé sous: {output_path}")
            
            return output_path
            
        except Exception as e:
            logger.error(f"Erreur: {e}")
            raise
        finally:
            if workbook:
                workbook.Close(SaveChanges=False)  # Ne pas sauvegarder l'original
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
    
    def _fill_data(self, sheet, df, metadata):
        """Remplit uniquement les données nécessaires"""
        # Métadonnées
        for row in range(1, 10):
            for col in range(1, 8):
                try:
                    cell_value = sheet.Cells(row, col).Value
                    if cell_value:
                        cell_text = str(cell_value).lower()
                        
                        if 'date de paiement' in cell_text:
                            sheet.Cells(row, col + 1).Value = metadata.get('date_paiement', '')
                        elif 'libellé' in cell_text or 'libelle' in cell_text:
                            sheet.Cells(row, col + 1).Value = metadata.get('libelle', '')
                        elif 'budget' in cell_text:
                            sheet.Cells(row, col + 1).Value = f"{metadata.get('budget', 500000):,}".replace(',', ' ')
                        elif 'projet' in cell_text:
                            sheet.Cells(row, col + 1).Value = metadata.get('projet', 'UGP')
                except:
                    pass
        
        # Trouver la ligne de début des transactions
        start_row = 9  # Ajuster selon votre template
        
        # Remplir les transactions
        for idx, record in df.iterrows():
            row = start_row + idx
            sheet.Cells(row, 1).Value = record.get('Date', '')
            sheet.Cells(row, 2).Value = record.get('TransactionID', '')
            sheet.Cells(row, 3).Value = 'PAIEMENT'
            sheet.Cells(row, 4).Value = record.get('Status', '')
            sheet.Cells(row, 5).Value = f"{record.get('Amount', 0):,.0f}".replace(',', ' ')
            sheet.Cells(row, 6).Value = f"{record.get('Frais', 0):,.0f}".replace(',', ' ')
            sheet.Cells(row, 7).Value = 'UGP'
            sheet.Cells(row, 8).Value = record.get('Vers', '')
            sheet.Cells(row, 9).Value = record.get('Beneficiaire', '')
