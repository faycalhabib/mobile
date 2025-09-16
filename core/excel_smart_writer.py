"""
Module intelligent pour écrire dans Excel avec insertion dynamique de lignes
Gère automatiquement l'espacement pour éviter l'écrasement des sections
"""
import win32com.client as win32
import pythoncom
import logging
import os
from datetime import datetime

logger = logging.getLogger(__name__)


class ExcelSmartWriter:
    """Écrit intelligemment dans Excel en ajustant les lignes"""
    
    def __init__(self):
        self.excel = None
        self.workbook = None
        self.sheet = None
        
        # Configuration du template
        self.HEADER_ROWS = 11  # Lignes 1-11: En-tête
        self.DATA_START_ROW = 12  # Première ligne de données
        self.MIN_DATA_ROWS = 2  # Minimum de lignes pour les données (12-13)
        self.RECAP_OFFSET = 3  # Lignes entre la fin des données et le récapitulatif
        self.RECAP_SECTION_ROWS = 6  # Nombre de lignes pour la section récapitulatif
        
    def open_excel(self, file_path):
        """Ouvre Excel et le fichier"""
        try:
            pythoncom.CoInitialize()
            self.excel = win32.Dispatch('Excel.Application')
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            
            self.workbook = self.excel.Workbooks.Open(os.path.abspath(file_path))
            self.sheet = self.workbook.Worksheets('Rapport paiement')
            
            logger.info(f"✓ Fichier Excel ouvert: {file_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Erreur ouverture Excel: {e}")
            return False
    
    def prepare_template(self, num_transactions):
        """
        Prépare le template en insérant des lignes si nécessaire
        
        Args:
            num_transactions: Nombre de transactions à écrire
        """
        logger.info(f"Préparation du template pour {num_transactions} transactions")
        
        # Calculer combien de lignes sont nécessaires
        needed_data_rows = num_transactions
        existing_data_rows = self.MIN_DATA_ROWS
        
        if needed_data_rows > existing_data_rows:
            # Il faut insérer des lignes
            rows_to_insert = needed_data_rows - existing_data_rows
            
            logger.info(f"  → Insertion de {rows_to_insert} lignes supplémentaires")
            
            # Point d'insertion: après la ligne 13
            insert_at_row = self.DATA_START_ROW + existing_data_rows
            
            # Insérer les lignes une par une
            for i in range(rows_to_insert):
                # Insérer une ligne
                self.sheet.Rows(insert_at_row).Insert()
                
                # Copier le format de la ligne 12 mais sans les bordures
                source_row = self.sheet.Rows(self.DATA_START_ROW)
                dest_row = self.sheet.Rows(insert_at_row)
                
                # Copier juste le format de police et alignement, pas les bordures
                for col in range(2, 11):  # Colonnes B à J
                    source_cell = self.sheet.Cells(self.DATA_START_ROW, col)
                    dest_cell = self.sheet.Cells(insert_at_row, col)
                    
                    # Copier le format
                    dest_cell.Font.Name = source_cell.Font.Name
                    dest_cell.Font.Size = source_cell.Font.Size
                    dest_cell.Font.Bold = False  # Pas de gras
                    dest_cell.HorizontalAlignment = -4108  # xlCenter
                    
                    # Pas de bordures
                    for border_idx in range(7, 13):  # Toutes les bordures
                        dest_cell.Borders(border_idx).LineStyle = -4142  # xlNone
                    
                    # Effacer le contenu
                    dest_cell.Value = ""
                
                logger.info(f"    • Ligne {insert_at_row} insérée (sans bordures)")
            
            # Nettoyer le presse-papier
            self.excel.CutCopyMode = False
            
        logger.info(f"  ✓ Template prêt avec {max(needed_data_rows, existing_data_rows)} lignes de données")
        
        return max(needed_data_rows, existing_data_rows)
    
    def write_metadata(self, metadata):
        """Écrit les métadonnées dans l'en-tête"""
        try:
            # Date de paiement (ligne 6, colonne C)
            if 'date_paiement' in metadata:
                self.sheet.Cells(6, 3).Value = metadata['date_paiement']
                logger.info(f"  • Date: {metadata['date_paiement']}")
            
            # Libellé (ligne 7, colonne C)
            if 'libelle' in metadata:
                self.sheet.Cells(7, 3).Value = metadata['libelle']
                logger.info(f"  • Libellé: {metadata['libelle']}")
            
            # Budget (ligne 8, colonne C)
            if 'budget' in metadata:
                # Formater le budget avec séparateurs
                budget_str = f"{int(metadata['budget']):,}".replace(',', ' ')
                self.sheet.Cells(8, 3).Value = budget_str
                logger.info(f"  • Budget: {budget_str}")
            
            # Projet (ligne 9, colonne C)
            if 'projet' in metadata:
                self.sheet.Cells(9, 3).Value = metadata['projet']
                logger.info(f"  • Projet: {metadata['projet']}")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Erreur écriture métadonnées: {e}")
            return False
    
    def write_transactions(self, df):
        """Écrit les transactions dans le tableau"""
        try:
            logger.info(f"\nÉcriture de {len(df)} transactions...")
            
            # Préparer le template pour le bon nombre de lignes
            total_data_rows = self.prepare_template(len(df))
            
            # Écrire chaque transaction
            for idx, row in df.iterrows():
                excel_row = self.DATA_START_ROW + idx
                
                # Date (colonne B)
                date_val = row.get('Date', '')
                if date_val:
                    cell = self.sheet.Cells(excel_row, 2)
                    cell.Value = str(date_val)
                    cell.HorizontalAlignment = -4108  # xlCenter
                    cell.Font.Bold = False  # Enlever le gras
                
                # N° Transaction (colonne C)
                trans_id = row.get('TransactionID', '')
                cell = self.sheet.Cells(excel_row, 3)
                cell.Value = str(trans_id)
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.Font.Bold = False
                
                # Type (colonne D)
                cell = self.sheet.Cells(excel_row, 4)
                cell.Value = 'PAIEMENT'
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.Font.Bold = False
                
                # Statut (colonne E)
                status = row.get('Status', 'Success')
                if status:
                    # Nettoyer le statut : enlever virgules et corriger l'orthographe
                    status = str(status).strip().replace('Succes', 'Success').replace(',', '')
                cell = self.sheet.Cells(excel_row, 5)
                cell.Value = status
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.Font.Bold = False
                
                # Montant (colonne F) - Formater avec espaces
                amount = row.get('Amount', 0)
                cell = self.sheet.Cells(excel_row, 6)
                if amount:
                    amount_str = f"{int(amount):,}".replace(',', ' ')
                    cell.Value = amount_str
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.Font.Bold = False
                
                # Frais ONG (colonne G)
                frais = row.get('Frais', 0)
                cell = self.sheet.Cells(excel_row, 7)
                if frais:
                    frais_str = f"{int(frais):,}".replace(',', ' ')
                    cell.Value = frais_str
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.Font.Bold = False
                
                # De (colonne H)
                cell = self.sheet.Cells(excel_row, 8)
                cell.Value = row.get('De', 'UGP')
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.Font.Bold = False
                
                # Vers (colonne I)
                vers = row.get('Vers', '')
                if vers:
                    # Enlever le préfixe du pays si présent
                    vers = str(vers).replace('235', '')
                cell = self.sheet.Cells(excel_row, 9)
                cell.Value = vers
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.Font.Bold = False
                
                # Bénéficiaire (colonne J)
                beneficiaire = row.get('Beneficiaire', '')
                cell = self.sheet.Cells(excel_row, 10)
                cell.Value = str(beneficiaire)
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.Font.Bold = False
                
                # Enlever les bordures inférieures de la ligne 13
                if excel_row == 13:
                    for col in range(2, 11):  # Colonnes B à J
                        cell = self.sheet.Cells(excel_row, col)
                        # Enlever la bordure inférieure
                        cell.Borders(9).LineStyle = -4142  # xlNone pour bordure inférieure
                
                logger.info(f"  ✓ Ligne {excel_row}: {trans_id} → {beneficiaire}")
            
            # Écrire le TOTAL
            self.write_total(df, total_data_rows)
            
            # Écrire le récapitulatif
            self.write_recapitulatif(df, total_data_rows)
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Erreur écriture transactions: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def write_total(self, df, total_data_rows):
        """Écrit la ligne de total"""
        try:
            # La ligne de total est 2 lignes après la dernière transaction
            total_row = self.DATA_START_ROW + total_data_rows + 1
            
            logger.info(f"\nÉcriture du TOTAL ligne {total_row}")
            
            # Texte "TOTAL:" en colonne E
            self.sheet.Cells(total_row, 5).Value = "TOTAL:"
            
            # Total des montants en colonne F
            total_amount = df['Amount'].sum()
            amount_str = f"{int(total_amount):,}".replace(',', ' ')
            self.sheet.Cells(total_row, 6).Value = amount_str
            
            # Total des frais en colonne G
            total_frais = df['Frais'].sum()
            frais_str = f"{int(total_frais):,}".replace(',', ' ')
            self.sheet.Cells(total_row, 7).Value = frais_str
            
            logger.info(f"  • Montant total: {amount_str} FCFA")
            logger.info(f"  • Frais totaux: {frais_str} FCFA")
            
            return total_row
            
        except Exception as e:
            logger.error(f"❌ Erreur écriture total: {e}")
            return 0
    
    def write_recapitulatif(self, df, total_data_rows):
        """Écrit SEULEMENT les valeurs dans la section récapitulatif existante"""
        try:
            # Le récapitulatif existe déjà dans le template, on cherche où il est
            # On va chercher le texte "Montant net à percevoir" pour trouver la bonne ligne
            
            logger.info(f"\nMise à jour des valeurs du récapitulatif")
            
            # Parcourir les lignes pour trouver le récapitulatif existant
            found_recap = False
            for row in range(20, 40):  # Chercher entre lignes 20 et 40
                cell_value = self.sheet.Cells(row, 1).Value
                if cell_value and "Montant net à percevoir" in str(cell_value):
                    found_recap = True
                    recap_row = row
                    
                    # Montant net à percevoir (déjà sur cette ligne)
                    total_amount = df['Amount'].sum()
                    amount_str = f"{int(total_amount):,}".replace(',', ' ')
                    self.sheet.Cells(recap_row, 10).Value = amount_str
                    logger.info(f"  • Montant net ligne {recap_row}: {amount_str}")
                    
                    # Frais (ligne suivante)
                    recap_row += 1
                    total_frais = df['Frais'].sum()
                    frais_str = f"{int(total_frais):,}".replace(',', ' ')
                    self.sheet.Cells(recap_row, 10).Value = frais_str
                    logger.info(f"  • Frais ligne {recap_row}: {frais_str}")
                    
                    # Total dépense (ligne suivante)
                    recap_row += 1
                    total_depense = total_amount + total_frais
                    depense_str = f"{int(total_depense):,}".replace(',', ' ')
                    self.sheet.Cells(recap_row, 10).Value = depense_str
                    logger.info(f"  • Total dépense ligne {recap_row}: {depense_str}")
                    
                    # Reliquat (ligne suivante)
                    recap_row += 1
                    # Le reliquat peut être calculé si on a le budget
                    # Pour l'instant, on laisse vide ou formule Excel
                    logger.info(f"  • Reliquat ligne {recap_row}: laissé tel quel")
                    
                    break
            
            if not found_recap:
                logger.warning("⚠️ Section récapitulatif non trouvée, création manuelle")
                # Si pas trouvé, créer à la position calculée
                recap_start_row = self.DATA_START_ROW + total_data_rows + self.RECAP_OFFSET
                
                # NE PAS écrire les labels, juste les valeurs
                # Les labels existent déjà dans le template
            
            logger.info(f"  ✓ Récapitulatif écrit")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Erreur écriture récapitulatif: {e}")
            return False
    
    def save_and_close(self):
        """Sauvegarde et ferme le fichier"""
        try:
            if self.workbook:
                self.workbook.Save()
                logger.info("✓ Fichier sauvegardé")
                
                self.workbook.Close(False)
            
            if self.excel:
                self.excel.Quit()
                
            pythoncom.CoUninitialize()
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Erreur fermeture: {e}")
            return False
    
    def write_report(self, file_path, df, metadata):
        """
        Méthode principale pour écrire le rapport complet
        
        Args:
            file_path: Chemin du fichier Excel
            df: DataFrame avec les transactions
            metadata: Dictionnaire avec les métadonnées
        """
        try:
            logger.info("\n" + "="*60)
            logger.info(" ÉCRITURE INTELLIGENTE DU RAPPORT EXCEL")
            logger.info("="*60)
            
            # Ouvrir Excel
            if not self.open_excel(file_path):
                return False
            
            # Écrire les métadonnées
            logger.info("\n📝 Écriture des métadonnées...")
            self.write_metadata(metadata)
            
            # Écrire les transactions
            logger.info("\n📊 Écriture des transactions...")
            self.write_transactions(df)
            
            # Sauvegarder et fermer
            logger.info("\n💾 Sauvegarde...")
            self.save_and_close()
            
            logger.info("\n✅ Rapport généré avec succès!")
            logger.info("="*60)
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Erreur génération rapport: {e}")
            import traceback
            traceback.print_exc()
            
            # Essayer de fermer proprement
            try:
                self.save_and_close()
            except:
                pass
                
            return False
