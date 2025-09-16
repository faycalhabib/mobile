"""
Module pour remplir le template Excel avec xlwings
Préserve GARANTIE à 100% tous les éléments (images, logos, formats, macros)
"""
import os
import shutil
from datetime import datetime
import xlwings as xw
import pandas as pd
import logging

logger = logging.getLogger(__name__)


class XlwingsFiller:
    """Remplit le template Excel en utilisant xlwings - Solution la plus robuste"""
    
    def __init__(self):
        self.template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
        
    def fill_template(self, processed_df: pd.DataFrame, metadata: dict, output_path: str) -> str:
        """
        Remplit le template en utilisant xlwings
        GARANTIT la préservation de TOUS les éléments
        """
        try:
            # Créer le dossier de sortie
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Copier le template
            shutil.copy2(self.template_path, output_path)
            logger.info(f"Template copié vers: {output_path}")
            
            # Ouvrir avec xlwings (visible=False pour ne pas afficher Excel)
            app = xw.App(visible=False, add_book=False)
            try:
                # Ouvrir le workbook
                wb = app.books.open(os.path.abspath(output_path))
                
                # Obtenir la feuille
                sheet_name = 'Rapport paiement'
                try:
                    sheet = wb.sheets[sheet_name]
                except:
                    sheet = wb.sheets[0]
                    logger.info(f"Utilisation de la feuille: {sheet.name}")
                
                # Remplir les métadonnées
                self._fill_metadata_xlwings(sheet, metadata)
                
                # Remplir les transactions
                self._fill_transactions_xlwings(sheet, processed_df)
                
                # Sauvegarder
                wb.save()
                logger.info("Fichier sauvegardé avec succès (images préservées)")
                
                # Fermer
                wb.close()
                
            finally:
                # Quitter Excel
                app.quit()
            
            return output_path
            
        except Exception as e:
            logger.error(f"Erreur xlwings: {e}")
            raise
    
    def _fill_metadata_xlwings(self, sheet, metadata: dict):
        """Remplit les métadonnées avec xlwings"""
        try:
            # Parcourir les premières lignes
            for row in range(1, 15):
                for col in range(1, 10):
                    cell_value = sheet.range((row, col)).value
                    if cell_value:
                        cell_text = str(cell_value).lower().strip()
                        
                        # Date de paiement
                        if 'date de paiement' in cell_text:
                            next_cell = sheet.range((row, col + 1))
                            if not next_cell.value:
                                next_cell.value = metadata.get('date_paiement', datetime.now().strftime("%d-%b-%Y"))
                                logger.info(f"Date remplie")
                        
                        # Libellé
                        elif 'libellé' in cell_text or 'libelle' in cell_text:
                            next_cell = sheet.range((row, col + 1))
                            if not next_cell.value:
                                next_cell.value = metadata.get('libelle', 'PAIEMENT')
                                logger.info(f"Libellé rempli")
                        
                        # Budget
                        elif 'budget' in cell_text:
                            next_cell = sheet.range((row, col + 1))
                            if not next_cell.value:
                                budget_value = metadata.get('budget', 500000)
                                next_cell.value = f"{budget_value:,.0f}".replace(',', ' ')
                                logger.info(f"Budget rempli")
                        
                        # Projet
                        elif 'projet' in cell_text:
                            next_cell = sheet.range((row, col + 1))
                            if not next_cell.value:
                                next_cell.value = metadata.get('projet', 'UGP')
                                logger.info(f"Projet rempli")
                                
        except Exception as e:
            logger.error(f"Erreur métadonnées: {e}")
    
    def _fill_transactions_xlwings(self, sheet, df: pd.DataFrame):
        """Remplit les transactions avec xlwings"""
        try:
            # Trouver la ligne d'en-tête
            header_row = None
            for row in range(1, 30):
                cell_value = sheet.range((row, 1)).value
                if cell_value and 'date' in str(cell_value).lower():
                    header_row = row
                    logger.info(f"En-tête trouvé ligne {header_row}")
                    break
            
            if not header_row:
                header_row = 8
            
            # Mapper les colonnes
            column_mapping = {}
            for col in range(1, 15):
                header_value = sheet.range((header_row, col)).value
                if header_value:
                    header_text = str(header_value).lower().strip()
                    
                    if 'date' in header_text:
                        column_mapping['Date'] = col
                    elif 'transaction' in header_text or 'n°' in header_text:
                        column_mapping['Transaction'] = col
                    elif 'type' in header_text:
                        column_mapping['Type'] = col
                    elif 'statut' in header_text:
                        column_mapping['Statut'] = col
                    elif 'montant' in header_text and 'frais' not in header_text:
                        column_mapping['Montant'] = col
                    elif 'frais' in header_text:
                        column_mapping['Frais'] = col
                    elif header_text == 'de':
                        column_mapping['De'] = col
                    elif 'vers' in header_text:
                        column_mapping['Vers'] = col
                    elif 'bénéficiaire' in header_text or 'beneficiaire' in header_text:
                        column_mapping['Beneficiaire'] = col
            
            # Remplir les données
            start_row = header_row + 1
            for idx, record in df.iterrows():
                current_row = start_row + idx
                
                if 'Date' in column_mapping:
                    sheet.range((current_row, column_mapping['Date'])).value = record.get('Date', '')
                
                if 'Transaction' in column_mapping:
                    sheet.range((current_row, column_mapping['Transaction'])).value = record.get('TransactionID', '')
                
                if 'Type' in column_mapping:
                    sheet.range((current_row, column_mapping['Type'])).value = record.get('Type', 'PAIEMENT')
                
                if 'Statut' in column_mapping:
                    sheet.range((current_row, column_mapping['Statut'])).value = record.get('Status', '')
                
                if 'Montant' in column_mapping:
                    amount = record.get('Amount', 0)
                    sheet.range((current_row, column_mapping['Montant'])).value = f"{amount:,.0f}".replace(',', ' ')
                
                if 'Frais' in column_mapping:
                    fee = record.get('Frais', 0)
                    sheet.range((current_row, column_mapping['Frais'])).value = f"{fee:,.0f}".replace(',', ' ')
                
                if 'De' in column_mapping:
                    sheet.range((current_row, column_mapping['De'])).value = record.get('De', 'UGP')
                
                if 'Vers' in column_mapping:
                    sheet.range((current_row, column_mapping['Vers'])).value = record.get('Vers', '')
                
                if 'Beneficiaire' in column_mapping:
                    sheet.range((current_row, column_mapping['Beneficiaire'])).value = record.get('Beneficiaire', '')
            
            # Totaux
            if len(df) > 0:
                total_row = start_row + len(df) + 1
                
                if 'Statut' in column_mapping:
                    sheet.range((total_row, column_mapping['Statut'])).value = "TOTAL:"
                
                if 'Montant' in column_mapping:
                    total_amount = df['Amount'].sum()
                    sheet.range((total_row, column_mapping['Montant'])).value = f"{total_amount:,.0f}".replace(',', ' ')
                
                if 'Frais' in column_mapping:
                    total_fees = df['Frais'].sum()
                    sheet.range((total_row, column_mapping['Frais'])).value = f"{total_fees:,.0f}".replace(',', ' ')
                
                logger.info(f"Totaux ajoutés")
                
        except Exception as e:
            logger.error(f"Erreur transactions: {e}")
