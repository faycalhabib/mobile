"""
Module pour remplir le template Excel via COM (Windows)
Garantit la préservation de TOUS les éléments (images, logos, formats)
"""
import os
import shutil
from datetime import datetime
import win32com.client as win32
import pythoncom
import pandas as pd
import logging

logger = logging.getLogger(__name__)


class ExcelCOMFiller:
    """Remplit le template Excel en utilisant Excel directement via COM"""
    
    def __init__(self):
        self.template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
        self.excel = None
        self.workbook = None
        
    def fill_template(self, processed_df: pd.DataFrame, metadata: dict, output_path: str) -> str:
        """
        Remplit le template en utilisant Excel via COM
        Préserve TOUT : images, logos, formats, formules, etc.
        """
        pythoncom.CoInitialize()
        
        try:
            # Créer le dossier de sortie
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Copier le template
            shutil.copy2(self.template_path, output_path)
            logger.info(f"Template copié vers: {output_path}")
            
            # Ouvrir Excel
            self.excel = win32.Dispatch('Excel.Application')
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            
            # Ouvrir le workbook
            self.workbook = self.excel.Workbooks.Open(os.path.abspath(output_path))
            logger.info("Fichier ouvert dans Excel")
            
            # Obtenir la feuille de travail
            sheet_name = 'Rapport paiement'
            try:
                sheet = self.workbook.Worksheets(sheet_name)
            except:
                sheet = self.workbook.Worksheets(1)  # Première feuille par défaut
                logger.info(f"Utilisation de la feuille: {sheet.Name}")
            
            # Remplir les métadonnées
            self._fill_metadata_com(sheet, metadata)
            
            # Remplir les transactions
            self._fill_transactions_com(sheet, processed_df)
            
            # Sauvegarder et fermer
            self.workbook.Save()
            logger.info("Fichier sauvegardé avec succès")
            
            return output_path
            
        except Exception as e:
            logger.error(f"Erreur COM: {e}")
            raise
            
        finally:
            # Nettoyer
            if self.workbook:
                self.workbook.Close(SaveChanges=True)
            if self.excel:
                self.excel.Quit()
            self.excel = None
            self.workbook = None
            pythoncom.CoUninitialize()
    
    def _fill_metadata_com(self, sheet, metadata: dict):
        """Remplit les métadonnées via COM"""
        try:
            # Parcourir les premières lignes pour trouver et remplir
            for row in range(1, 15):
                for col in range(1, 10):
                    cell_value = sheet.Cells(row, col).Value
                    if cell_value:
                        cell_text = str(cell_value).lower().strip()
                        
                        # Date de paiement
                        if 'date de paiement' in cell_text:
                            next_cell = sheet.Cells(row, col + 1)
                            if not next_cell.Value:
                                next_cell.Value = metadata.get('date_paiement', datetime.now().strftime("%d-%b-%Y"))
                                logger.info(f"Date remplie en ligne {row}")
                        
                        # Libellé
                        elif 'libellé' in cell_text or 'libelle' in cell_text:
                            next_cell = sheet.Cells(row, col + 1)
                            if not next_cell.Value:
                                next_cell.Value = metadata.get('libelle', 'PAIEMENT')
                                logger.info(f"Libellé rempli en ligne {row}")
                        
                        # Budget
                        elif 'budget' in cell_text:
                            next_cell = sheet.Cells(row, col + 1)
                            if not next_cell.Value:
                                budget_value = metadata.get('budget', 500000)
                                next_cell.Value = f"{budget_value:,.0f}".replace(',', ' ')
                                logger.info(f"Budget rempli en ligne {row}")
                        
                        # Projet
                        elif 'projet' in cell_text:
                            next_cell = sheet.Cells(row, col + 1)
                            if not next_cell.Value:
                                next_cell.Value = metadata.get('projet', 'UGP')
                                logger.info(f"Projet rempli en ligne {row}")
                                
        except Exception as e:
            logger.error(f"Erreur remplissage métadonnées: {e}")
    
    def _fill_transactions_com(self, sheet, df: pd.DataFrame):
        """Remplit les transactions via COM"""
        try:
            # Trouver la ligne d'en-tête
            header_row = None
            for row in range(1, 30):
                cell_value = sheet.Cells(row, 1).Value
                if cell_value and 'date' in str(cell_value).lower():
                    # Vérifier d'autres colonnes pour confirmer
                    other_cell = sheet.Cells(row, 2).Value
                    if other_cell and any(x in str(other_cell).lower() for x in ['transaction', 'n°']):
                        header_row = row
                        logger.info(f"En-tête trouvé à la ligne {header_row}")
                        break
            
            if not header_row:
                header_row = 8  # Valeur par défaut
                logger.warning(f"En-tête non trouvé, utilisation ligne {header_row}")
            
            # Mapper les colonnes
            column_mapping = {}
            for col in range(1, 15):
                header_value = sheet.Cells(header_row, col).Value
                if header_value:
                    header_text = str(header_value).lower().strip()
                    
                    if 'date' in header_text:
                        column_mapping['Date'] = col
                    elif 'transaction' in header_text or 'n°' in header_text:
                        column_mapping['Transaction'] = col
                    elif 'type' in header_text:
                        column_mapping['Type'] = col
                    elif 'statut' in header_text or 'status' in header_text:
                        column_mapping['Statut'] = col
                    elif 'montant' in header_text and 'frais' not in header_text:
                        column_mapping['Montant'] = col
                    elif 'frais' in header_text:
                        column_mapping['Frais'] = col
                    elif header_text == 'de':
                        column_mapping['De'] = col
                    elif 'vers' in header_text or 'à' in header_text:
                        column_mapping['Vers'] = col
                    elif 'bénéficiaire' in header_text or 'beneficiaire' in header_text:
                        column_mapping['Beneficiaire'] = col
            
            logger.info(f"Colonnes mappées: {column_mapping}")
            
            # Remplir les données
            start_row = header_row + 1
            for idx, record in df.iterrows():
                current_row = start_row + idx
                
                if 'Date' in column_mapping:
                    sheet.Cells(current_row, column_mapping['Date']).Value = record.get('Date', '')
                
                if 'Transaction' in column_mapping:
                    sheet.Cells(current_row, column_mapping['Transaction']).Value = record.get('TransactionID', '')
                
                if 'Type' in column_mapping:
                    sheet.Cells(current_row, column_mapping['Type']).Value = record.get('Type', 'PAIEMENT')
                
                if 'Statut' in column_mapping:
                    sheet.Cells(current_row, column_mapping['Statut']).Value = record.get('Status', '')
                
                if 'Montant' in column_mapping:
                    amount = record.get('Amount', 0)
                    sheet.Cells(current_row, column_mapping['Montant']).Value = f"{amount:,.0f}".replace(',', ' ')
                
                if 'Frais' in column_mapping:
                    fee = record.get('Frais', 0)
                    sheet.Cells(current_row, column_mapping['Frais']).Value = f"{fee:,.0f}".replace(',', ' ')
                
                if 'De' in column_mapping:
                    sheet.Cells(current_row, column_mapping['De']).Value = record.get('De', 'UGP')
                
                if 'Vers' in column_mapping:
                    sheet.Cells(current_row, column_mapping['Vers']).Value = record.get('Vers', '')
                
                if 'Beneficiaire' in column_mapping:
                    sheet.Cells(current_row, column_mapping['Beneficiaire']).Value = record.get('Beneficiaire', '')
            
            # Ajouter les totaux
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
                
                logger.info(f"Totaux ajoutés ligne {total_row}")
                
        except Exception as e:
            logger.error(f"Erreur remplissage transactions: {e}")
