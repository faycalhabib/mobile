"""
Module de génération de rapports Excel
"""
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional
import pandas as pd
import logging

logger = logging.getLogger(__name__)


class ReportGenerator:
    """Générateur de rapports Excel formatés"""
    
    def __init__(self, config: dict = None):
        self.config = config or {}
        self.output_dir = self.config.get('preferences', {}).get('output_folder', './outputs')
        
    def generate_report(self, data: pd.DataFrame, metadata: dict, output_name: str = None) -> str:
        """
        Génère le rapport Excel à partir des données
        
        Args:
            data: DataFrame avec les transactions
            metadata: Métadonnées du rapport
            output_name: Nom du fichier de sortie (optionnel)
            
        Returns:
            Chemin du fichier généré
        """
        if output_name is None:
            output_name = f"Rapport_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        output_path = Path(self.config['preferences']['output_folder']) / output_name
        
        try:
            # Vérifier si on doit utiliser le mode rapide
            use_fast_mode = self.config.get('optimization', {}).get('use_fast_mode', False)
            
            if use_fast_mode:
                logger.info("🚀 Utilisation du FastWriter optimisé")
                try:
                    from core.excel_fast_writer import ExcelHybridWriter
                    template_path = Path(__file__).parent.parent / 'templates' / 'Rapport_template.xlsx'
                    
                    # Créer le dossier de sortie si nécessaire
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    
                    # Utiliser le writer rapide
                    fast_writer = ExcelHybridWriter(
                        template_path=str(template_path),
                        output_path=str(output_path)
                    )
                    return fast_writer.write_report(data, metadata)
                    
                except Exception as e:
                    logger.warning(f"⚠ FastWriter échoué, fallback au mode classique: {e}")
                    # Fallback au mode classique
                    
            # Mode classique (par défaut ou si fast mode échoue)
            logger.info("Utilisation de FinalExcelFiller avec win32com")
            from core.final_excel_filler import FinalExcelFiller
            template_path = Path(__file__).parent.parent / 'templates' / 'Rapport_template.xlsx'
            
            # Créer le dossier de sortie si nécessaire
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            filler = FinalExcelFiller()
            # Ordre correct des arguments: template_path, output_path, df, metadata
            success = filler.fill_template(str(template_path), str(output_path), data, metadata)
            
            if success:
                return str(output_path)
            else:
                return None
        
        except Exception as e:
            logger.error(f"Erreur lors de la génération du rapport: {e}")
            return None

        
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
