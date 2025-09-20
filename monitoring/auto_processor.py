"""
Processeur automatique principal qui orchestre tout le workflow
"""
import os
import sys
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional

# Ajouter le dossier parent au path
sys.path.append(str(Path(__file__).parent.parent))

from core.file_handler import FileHandler
from core.data_processor import DataProcessor
from core.report_generator import ReportGenerator
from monitoring.pdf_converter import ProfessionalPDFConverter
from monitoring.email_sender import ProfessionalEmailSender
from monitoring.file_watcher_fixed import SmartFileWatcher
import json

# Configuration du logging avec support UTF-8
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/auto_processor.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


class AutoProcessor:
    """Orchestrateur principal du traitement automatique"""
    
    def __init__(self, config_path: str = "config/auto_processor_config.json"):
        """
        Initialise le processeur automatique
        
        Args:
            config_path: Chemin vers la configuration
        """
        self.config = self._load_config(config_path)
        
        # Initialiser les composants
        self.file_handler = FileHandler()
        self.data_processor = DataProcessor()
        self.data_processor.use_smart_processing = True
        self.report_generator = ReportGenerator(self.config)
        self.pdf_converter = ProfessionalPDFConverter()
        self.email_sender = ProfessionalEmailSender()
        self.file_watcher = SmartFileWatcher()
        
        # Statistiques de traitement
        self.processing_stats = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'last_process': None
        }
        
        logger.info("🚀 AutoProcessor initialisé et prêt")
    
    def _load_config(self, config_path: str) -> dict:
        """Charge la configuration du processeur"""
        default_config = {
            'preferences': {
                'output_folder': './outputs'
            },
            'processing': {
                'auto_retry': True,
                'max_retries': 3,
                'cleanup_after': True,
                'generate_pdf': True,
                'send_email': True
            },
            'metadata': {
                'date_paiement': datetime.now().strftime('%d/%m/%Y'),
                'libelle': 'PAIEMENT AUTOMATIQUE',
                'budget': 5000000,
                'projet': 'UGP'
            }
        }
        
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    self._merge_configs(default_config, user_config)
        except Exception as e:
            logger.warning(f"⚠️ Config par défaut utilisée: {e}")
        
        return default_config
    
    def _merge_configs(self, default: dict, user: dict):
        """Fusionne les configurations"""
        for key, value in user.items():
            if key in default and isinstance(default[key], dict) and isinstance(value, dict):
                self._merge_configs(default[key], value)
            else:
                default[key] = value
    
    def process_files(self, files: Dict[str, str]) -> Dict:
        """
        Traite un ensemble de fichiers complet
        
        Args:
            files: Dictionnaire avec les chemins des fichiers
                  {'bulkreport': ..., 'export': ..., 'frais': ...}
        
        Returns:
            Dictionnaire avec le résultat du traitement
        """
        logger.info("="*70)
        logger.info(" DÉBUT DU TRAITEMENT AUTOMATIQUE")
        logger.info("="*70)
        
        result = {
            'success': False,
            'error': None,
            'report_path': None,
            'pdf_path': None,
            'email_sent': False,
            'timestamp': datetime.now(),
            'stats': {}
        }
        
        try:
            # 1. LECTURE DES FICHIERS
            logger.info("\n📁 ÉTAPE 1: Lecture des fichiers")
            
            bulk_df, metadata = self.file_handler.read_bulk_report(files['bulkreport'])
            logger.info(f"  ✓ BulkReport: {len(bulk_df)} transactions")
            
            export_df = self.file_handler.read_export_file(files['export'])
            logger.info(f"  ✓ Export: {len(export_df)} bénéficiaires")
            
            fees_df = None
            if files.get('frais') and files['frais']:
                try:
                    fees_df = self.file_handler.read_fees_file(files['frais'])
                    logger.info(f"  ✓ Frais: Table chargée")
                except:
                    logger.warning(f"  ⚠ Frais: Utilisation du taux par défaut")
            
            # 2. TRAITEMENT DES DONNÉES
            logger.info("\n🔄 ÉTAPE 2: Traitement des données")
            
            processed_df, errors = self.data_processor.process_transactions(
                bulk_df, export_df, fees_df if fees_df is not None else None,
                self.config['metadata']
            )
            
            logger.info(f"  ✓ {len(processed_df)} transactions traitées")
            if errors:
                logger.warning(f"  ⚠ {len(errors)} avertissements")
            
            # 3. GÉNÉRATION DU RAPPORT EXCEL
            logger.info("\n📊 ÉTAPE 3: Génération du rapport Excel")
            
            report_name = f"Rapport_AUTO_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            report_path = self.report_generator.generate_report(
                processed_df,
                self.config['metadata'],
                report_name
            )
            
            if report_path is None:
                raise Exception("Échec de la génération du rapport Excel")
                
            result['report_path'] = report_path
            logger.info(f"  ✓ Rapport généré: {Path(report_path).name}")
            
            # Statistiques pour email
            result['stats'] = {
                'transaction_count': len(processed_df),
                'total_amount': processed_df['Amount'].sum() if 'Amount' in processed_df.columns else 0,
                'total_fees': processed_df['Frais'].sum() if 'Frais' in processed_df.columns else 0,
                'unique_beneficiaries': processed_df['Beneficiaire'].nunique() if 'Beneficiaire' in processed_df.columns else 0,
                'date': self.config['metadata']['date_paiement']
            }
            
            # 4. CONVERSION EN PDF
            if self.config['processing']['generate_pdf']:
                logger.info("\n📄 ÉTAPE 4: Conversion en PDF")
                
                pdf_result = self.pdf_converter.convert_excel_to_pdf(
                    report_path,
                    options={
                        'quality': 'standard',
                        'orientation': 'portrait',
                        'fit_to_page': True,
                        'center_horizontally': True
                    }
                )
                
                if pdf_result['success']:
                    result['pdf_path'] = pdf_result['pdf_path']
                    logger.info(f"  ✓ PDF généré: {Path(pdf_result['pdf_path']).name}")
                else:
                    logger.warning(f"  ⚠ Échec conversion PDF: {pdf_result.get('error')}")
            
            # 5. ENVOI PAR EMAIL
            if self.config['processing']['send_email'] and result['pdf_path']:
                logger.info("\n📧 ÉTAPE 5: Envoi par email")
                
                attachments = []
                if result['pdf_path']:
                    attachments.append(result['pdf_path'])
                # Optionnel: ajouter aussi l'Excel
                # attachments.append(result['report_path'])
                
                email_results = self.email_sender.send_to_all_partners(
                    result['stats'],
                    attachments
                )
                
                result['email_sent'] = len(email_results['success']) > 0
                
                if result['email_sent']:
                    logger.info(f"  ✓ Emails envoyés: {len(email_results['success'])} destinataires")
                else:
                    logger.warning(f"  ⚠ Échec envoi emails")
            
            # Succès global
            result['success'] = True
            self.processing_stats['success'] += 1
            logger.info("\n✅ TRAITEMENT TERMINÉ AVEC SUCCÈS")
            
        except Exception as e:
            logger.error(f"\n❌ ERREUR: {str(e)}")
            result['error'] = str(e)
            self.processing_stats['failed'] += 1
            
            # Tentative de notification d'erreur
            try:
                # Envoyer email d'erreur aux admins
                pass
            except:
                pass
        
        finally:
            self.processing_stats['total'] += 1
            self.processing_stats['last_process'] = datetime.now()
        
        # Afficher le résumé
        self._log_summary(result)
        
        return result
    
    def _log_summary(self, result: Dict):
        """Affiche un résumé du traitement"""
        logger.info("\n" + "="*70)
        logger.info(" RÉSUMÉ DU TRAITEMENT")
        logger.info("="*70)
        
        if result['success']:
            logger.info(f"✅ Statut: SUCCÈS")
            logger.info(f"📊 Rapport: {Path(result['report_path']).name if result['report_path'] else 'N/A'}")
            logger.info(f"📄 PDF: {Path(result['pdf_path']).name if result['pdf_path'] else 'N/A'}")
            logger.info(f"📧 Email: {'Envoyé' if result['email_sent'] else 'Non envoyé'}")
            
            if result['stats']:
                logger.info(f"\n📈 Statistiques:")
                logger.info(f"  • Transactions: {result['stats']['transaction_count']}")
                logger.info(f"  • Montant total: {result['stats']['total_amount']:,.0f} FCFA")
                logger.info(f"  • Frais totaux: {result['stats']['total_fees']:,.0f} FCFA")
        else:
            logger.info(f"❌ Statut: ÉCHEC")
            logger.info(f"🔴 Erreur: {result['error']}")
        
        logger.info("="*70)
    
    def start_monitoring(self):
        """Démarre le monitoring automatique du dossier"""
        logger.info("\n🚀 DÉMARRAGE DU MONITORING AUTOMATIQUE")
        logger.info("="*70)
        
        # Configurer le callback
        self.file_watcher.set_process_callback(self.process_files)
        
        # Démarrer le monitoring
        logger.info(f"👁️ Surveillance du dossier: {self.file_watcher.watched_folder}")
        logger.info("  → Déposez les fichiers requis pour déclencher le traitement")
        logger.info("  → BulkReport.csv + Export.xlsx (+ Frais.xlsx optionnel)")
        logger.info("  → Ctrl+C pour arrêter\n")
        
        self.file_watcher.start_monitoring()
    
    def get_stats(self) -> Dict:
        """Retourne les statistiques globales"""
        return {
            'processor': self.processing_stats,
            'pdf_converter': self.pdf_converter.get_stats(),
            'email_sender': self.email_sender.get_stats(),
            'file_watcher': self.file_watcher.get_stats()
        }


def main():
    """Point d'entrée principal"""
    print("""
    ╔══════════════════════════════════════════════════════════╗
    ║                                                          ║
    ║            UGP REPORTER - MODE AUTOMATIQUE              ║
    ║                                                          ║
    ║  🚀 Monitoring intelligent des dossiers                 ║
    ║  📊 Génération automatique de rapports                  ║
    ║  📄 Conversion PDF professionnelle                      ║
    ║  📧 Envoi email aux partenaires                         ║
    ║                                                          ║
    ╚══════════════════════════════════════════════════════════╝
    """)
    
    try:
        processor = AutoProcessor()
        processor.start_monitoring()
    except KeyboardInterrupt:
        print("\n\n⏹️ Arrêt du monitoring...")
        print("👋 Au revoir!")
    except Exception as e:
        print(f"\n❌ Erreur fatale: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
