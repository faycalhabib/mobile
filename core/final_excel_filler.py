"""
Solution FINALE : Préserve les 531 images/formes du template
"""
import os
import win32com.client as win32
import pythoncom
import pandas as pd
import logging
from datetime import datetime

logger = logging.getLogger(__name__)


class FinalExcelFiller:
    """Solution définitive qui préserve TOUTES les images"""
    
    def __init__(self):
        self.template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
        
    def fill_template(self, template_path: str, output_path: str, df: pd.DataFrame, metadata: dict) -> bool:
        """
        Remplit le template Excel avec insertion intelligente de lignes
        """
        import shutil
        import os
        
        # Copier d'abord le template vers le fichier de sortie
        try:
            # Créer le répertoire de sortie si nécessaire
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Copier le template
            shutil.copy2(template_path, output_path)
            logger.info(f"Template copié vers: {output_path}")
            
        except Exception as e:
            logger.error(f"Erreur copie template: {e}")
            return False
        
        # Utiliser le nouveau writer intelligent
        from .excel_smart_writer import ExcelSmartWriter
        
        logger.info("Utilisation du ExcelSmartWriter pour l'écriture intelligente")
        
        writer = ExcelSmartWriter()
        success = writer.write_report(output_path, df, metadata)
        
        return success
    
    def fill_template_old(self, processed_df: pd.DataFrame, metadata: dict, output_path: str) -> str:
        """
        Ouvre le template, remplit les données, SaveAs avec FileFormat approprié
        """
        pythoncom.CoInitialize()
        excel = None
        workbook = None
        
        try:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Ouvrir Excel
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            
            # Ouvrir le template ORIGINAL
            workbook = excel.Workbooks.Open(os.path.abspath(self.template_path), ReadOnly=False)
            logger.info(f"Template ouvert: {self.template_path}")
            
            # Obtenir la feuille 'Rapport paiement'
            sheet = workbook.Worksheets('Rapport paiement')
            
            # Vérifier le nombre d'images avant
            shapes_before = sheet.Shapes.Count
            logger.info(f"Images/formes dans le template: {shapes_before}")
            
            # Remplir les données SANS toucher aux images
            self._fill_metadata(sheet, metadata)
            self._fill_transactions(sheet, processed_df)
            
            # IMPORTANT: SaveAs avec le bon FileFormat pour préserver les images
            # 51 = xlOpenXMLWorkbook (xlsx avec toutes les fonctionnalités)
            workbook.SaveAs(os.path.abspath(output_path), FileFormat=51)
            
            # Vérifier le nombre d'images après
            shapes_after = sheet.Shapes.Count
            logger.info(f"Images/formes après sauvegarde: {shapes_after}")
            
            if shapes_after != shapes_before:
                logger.warning(f"Attention: {shapes_before - shapes_after} images perdues")
            else:
                logger.info("✓ Toutes les images préservées!")
            
            return output_path
            
        except Exception as e:
            logger.error(f"Erreur: {e}")
            raise
        finally:
            if workbook:
                workbook.Close(SaveChanges=False)
            if excel:
                excel.ScreenUpdating = True
                excel.Quit()
            pythoncom.CoUninitialize()
    
    def _fill_metadata(self, sheet, metadata):
        """Remplit les métadonnées sans toucher aux images"""
        # Recherche ciblée dans les premières lignes
        for row in range(1, 10):
            cell_b = sheet.Cells(row, 2).Value  # Colonne B
            if cell_b and 'date de paiement' in str(cell_b).lower():
                sheet.Cells(row, 3).Value = metadata.get('date_paiement', datetime.now().strftime("%d/%m/%Y"))
            elif cell_b and 'libellé' in str(cell_b).lower():
                sheet.Cells(row, 3).Value = metadata.get('libelle', 'PAIEMENT')
            elif cell_b and 'budget' in str(cell_b).lower():
                sheet.Cells(row, 3).Value = metadata.get('budget', 500000)
            elif cell_b and 'projet' in str(cell_b).lower():
                sheet.Cells(row, 3).Value = metadata.get('projet', 'UGP')
    
    def _fill_transactions(self, sheet, df):
        """Remplit les transactions sans toucher aux images"""
        # L'en-tête du tableau est à la ligne 11 (confirmé par le debug)
        header_row = 11
        logger.info(f"En-tête du tableau à la ligne {header_row}")
        
        # Mapper les colonnes selon l'analyse (colonnes 2 à 10)
        # D'après le debug: Col 2: Date, Col 3: N° Transaction, Col 4: Type, 
        # Col 5: Statut, Col 6: Montant, Col 7: Frais ONG, Col 8: De, Col 9: Vers, Col 10: Beneficiaire
        column_mapping = {
            'Date': 2,
            'Transaction': 3,
            'Type': 4,
            'Statut': 5,
            'Montant': 6,
            'Frais': 7,
            'De': 8,
            'Vers': 9,
            'Beneficiaire': 10
        }
        
        logger.info(f"Colonnes mappées: {column_mapping}")
        
        # Remplir les données à partir de la ligne suivante
        start_row = header_row + 1
        
        logger.info(f"Début remplissage de {len(df)} transactions à partir de ligne {start_row}")
        
        for idx, record in df.iterrows():
            row = start_row + idx
            
            # Log détaillé avant écriture
            logger.info(f"Écriture ligne {row}:")
            logger.info(f"  - Date: {record.get('Date', '')}")
            logger.info(f"  - TransactionID: {record.get('TransactionID', '')}")
            logger.info(f"  - Amount: {record.get('Amount', 0)}")
            
            try:
                # Écrire directement dans les colonnes sans vérification (colonnes 2-10)
                sheet.Cells(row, 2).Value = str(record.get('Date', ''))  # Date
                sheet.Cells(row, 3).Value = str(record.get('TransactionID', ''))  # N° Transaction
                sheet.Cells(row, 4).Value = 'PAIEMENT'  # Type
                sheet.Cells(row, 5).Value = str(record.get('Status', 'Success')).strip()  # Statut
                
                amount = record.get('Amount', 0)
                sheet.Cells(row, 6).Value = f"{amount:,.0f}".replace(',', ' ')  # Montant
                
                fee = record.get('Frais', 0)
                sheet.Cells(row, 7).Value = f"{fee:,.0f}".replace(',', ' ')  # Frais ONG
                
                sheet.Cells(row, 8).Value = 'UGP'  # De
                sheet.Cells(row, 9).Value = str(record.get('Vers', ''))  # Vers
                sheet.Cells(row, 10).Value = str(record.get('Beneficiaire', ''))  # Bénéficiaire
                
                logger.info(f"✓ Ligne {row} écrite avec succès")
            except Exception as e:
                logger.error(f"✗ Erreur écriture ligne {row}: {e}")
        
        # Ajouter la ligne de total
        if len(df) > 0:
            total_row = start_row + len(df) + 1
            
            if 'Statut' in column_mapping:
                sheet.Cells(total_row, column_mapping['Statut']).Value = "TOTAL:"
            
            if 'Montant' in column_mapping:
                total_amount = df['Amount'].sum()
                sheet.Cells(total_row, column_mapping['Montant']).Value = f"{total_amount:,.0f}".replace(',', ' ')
            
            if 'Frais' in column_mapping:
                total_fees = df['Frais'].sum()
                sheet.Cells(total_row, column_mapping['Frais']).Value = f"{total_fees:,.0f}".replace(',', ' ')
            
        logger.info(f"Rempli {len(df)} transactions")
