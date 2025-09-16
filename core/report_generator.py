"""
Module de génération de rapports Excel
"""
import os
from datetime import datetime
from typing import Dict, Optional
import pandas as pd
import logging

logger = logging.getLogger(__name__)


class ReportGenerator:
    """Générateur de rapports Excel formatés"""
    
    def __init__(self, config: dict = None):
        self.config = config or {}
        self.output_dir = self.config.get('preferences', {}).get('output_folder', './outputs')
        # Ajouter le chemin du template
        self.template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
        
    def generate_report(self, 
                       processed_df: pd.DataFrame,
                       metadata: dict,
                       output_name: Optional[str] = None) -> str:
        """
        Générer le rapport Excel final
        
        Args:
            processed_df: DataFrame avec les données traitées
            metadata: Métadonnées du rapport (date, libellé, budget, etc.)
            output_name: Nom du fichier de sortie (optionnel)
            
        Returns:
            Chemin du fichier généré
        """
        # Créer le nom du fichier si non fourni
        if not output_name:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"Rapport_UGP_{timestamp}.xlsx"
        
        # Chemin complet
        output_path = os.path.join(self.output_dir, output_name)
        
        # Créer le dossier de sortie
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Sauvegarder avec le file_handler pour un formatage approprié
        # FORCER l'utilisation de FinalExcelFiller qui fonctionne
        from core.final_excel_filler import FinalExcelFiller
        filler = FinalExcelFiller()
        logger.info("Utilisation de FinalExcelFiller avec win32com")
        
        # Ajouter les métadonnées complètes
        full_metadata = {
            'date_paiement': metadata.get('date_paiement', datetime.now().strftime("%d-%b-%Y")),
            'libelle': metadata.get('libelle', 'PAIEMENT'),
            'budget': metadata.get('budget', 500000),
            'projet': metadata.get('projet', 'UGP')
        }
        
        # Générer le rapport en utilisant le template existant
        # La nouvelle signature attend: template_path, output_path, df, metadata
        output_file = filler.fill_template(self.template_path, output_path, processed_df, full_metadata)
        
        # Si fill_template retourne un booléen, retourner le chemin de sortie
        if output_file == True:
            output_file = output_path
        
        logger.info(f"Rapport généré avec succès: {output_file}")
        
        return output_file
    
    def create_summary_sheet(self, writer, stats: dict, errors: list):
        """Créer une feuille de résumé dans le rapport Excel"""
        summary_data = {
            'Statistique': [
                'Nombre de transactions',
                'Montant total',
                'Frais totaux',
                'Nombre de bénéficiaires',
                'Montant moyen',
                'Montant minimum',
                'Montant maximum',
                '',
                'Avertissements:'
            ],
            'Valeur': [
                stats['total_transactions'],
                f"{stats['total_amount']:,.0f} FCFA",
                f"{stats['total_fees']:,.0f} FCFA",
                stats['unique_beneficiaries'],
                f"{stats['average_amount']:,.0f} FCFA",
                f"{stats['min_amount']:,.0f} FCFA",
                f"{stats['max_amount']:,.0f} FCFA",
                '',
                ''
            ]
        }
        
        # Ajouter les erreurs
        for error in errors[:10]:  # Limiter à 10 erreurs
            summary_data['Statistique'].append('')
            summary_data['Valeur'].append(error)
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Résumé', index=False)
