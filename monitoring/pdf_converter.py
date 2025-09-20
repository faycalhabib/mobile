"""
Convertisseur Excel vers PDF professionnel avec mise en page optimisée
"""
import os
import logging
from pathlib import Path
import win32com.client as win32
import pythoncom
from typing import Optional, Dict
import time

logger = logging.getLogger(__name__)


class ProfessionalPDFConverter:
    """Convertisseur Excel vers PDF avec options avancées"""
    
    def __init__(self):
        self.excel = None
        self.conversion_stats = {
            'total': 0,
            'success': 0,
            'failed': 0
        }
    
    def convert_excel_to_pdf(self, excel_path: str, pdf_path: Optional[str] = None,
                            options: Optional[Dict] = None) -> Dict:
        """
        Convertit un fichier Excel en PDF avec mise en page professionnelle
        
        Args:
            excel_path: Chemin du fichier Excel
            pdf_path: Chemin de sortie PDF (optionnel)
            options: Options de conversion
        
        Returns:
            Dict avec statut et chemin du PDF
        """
        pythoncom.CoInitialize()
        
        try:
            # Configuration par défaut
            default_options = {
                'quality': 'standard',  # standard, minimum, maximum
                'orientation': 'portrait',  # portrait, landscape
                'fit_to_page': True,
                'margins': 'normal',  # normal, narrow, wide
                'include_headers': True,
                'center_horizontally': True,
                'center_vertically': False,
                'grid_lines': False
            }
            
            if options:
                default_options.update(options)
            
            # Générer le chemin PDF si non fourni
            if not pdf_path:
                excel_path_obj = Path(excel_path)
                pdf_path = excel_path_obj.parent / f"{excel_path_obj.stem}_report.pdf"
            
            logger.info(f"📄 Conversion Excel → PDF")
            logger.info(f"  Source: {Path(excel_path).name}")
            logger.info(f"  Destination: {Path(pdf_path).name}")
            
            # Ouvrir Excel
            self.excel = win32.Dispatch('Excel.Application')
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            
            # Ouvrir le fichier
            workbook = self.excel.Workbooks.Open(os.path.abspath(excel_path))
            
            # Configurer la mise en page
            for sheet in workbook.Worksheets:
                self._configure_page_setup(sheet, default_options)
            
            # Sélectionner la feuille principale
            main_sheet = workbook.Worksheets('Rapport paiement')
            main_sheet.Select()
            
            # Exporter en PDF avec qualité maximale
            export_params = self._get_export_params(default_options)
            
            workbook.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=os.path.abspath(pdf_path),
                Quality=export_params['quality'],
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            
            # Fermer le workbook
            workbook.Close(SaveChanges=False)
            
            # Statistiques
            self.conversion_stats['total'] += 1
            self.conversion_stats['success'] += 1
            
            # Vérifier que le PDF a été créé
            if os.path.exists(pdf_path):
                file_size = os.path.getsize(pdf_path)
                logger.info(f"✅ PDF généré avec succès ({self._format_size(file_size)})")
                
                return {
                    'success': True,
                    'pdf_path': str(pdf_path),
                    'file_size': file_size,
                    'timestamp': time.time()
                }
            else:
                raise Exception("Le fichier PDF n'a pas été créé")
            
        except Exception as e:
            logger.error(f"❌ Erreur conversion PDF: {e}")
            self.conversion_stats['failed'] += 1
            
            return {
                'success': False,
                'error': str(e),
                'timestamp': time.time()
            }
        
        finally:
            # Nettoyer Excel
            if self.excel:
                try:
                    self.excel.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
    
    def _configure_page_setup(self, sheet, options: Dict):
        """Configure la mise en page pour l'impression"""
        try:
            ps = sheet.PageSetup
            
            # Orientation
            if options['orientation'] == 'landscape':
                ps.Orientation = 2  # xlLandscape
            else:
                ps.Orientation = 1  # xlPortrait
            
            # Marges (en pouces)
            margins_config = {
                'narrow': {'top': 0.5, 'bottom': 0.5, 'left': 0.5, 'right': 0.5},
                'normal': {'top': 0.75, 'bottom': 0.75, 'left': 0.7, 'right': 0.7},
                'wide': {'top': 1, 'bottom': 1, 'left': 1, 'right': 1}
            }
            
            margins = margins_config.get(options['margins'], margins_config['normal'])
            ps.TopMargin = self.excel.InchesToPoints(margins['top'])
            ps.BottomMargin = self.excel.InchesToPoints(margins['bottom'])
            ps.LeftMargin = self.excel.InchesToPoints(margins['left'])
            ps.RightMargin = self.excel.InchesToPoints(margins['right'])
            
            # Ajuster à la page
            if options['fit_to_page']:
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False  # Automatique
            
            # Centrage
            ps.CenterHorizontally = options['center_horizontally']
            ps.CenterVertically = options['center_vertically']
            
            # En-têtes et pieds de page
            if options['include_headers']:
                ps.LeftHeader = "&D"  # Date
                ps.CenterHeader = "&A"  # Nom de la feuille
                ps.RightHeader = "&P/&N"  # Page X sur Y
                ps.CenterFooter = "UGP Reporter - Rapport Automatique"
            
            # Quadrillage
            ps.PrintGridlines = options['grid_lines']
            
            # Zone d'impression (automatique basée sur les données)
            ps.PrintArea = ""  # Reset pour auto-detect
            
            # Ordre des pages
            ps.Order = 1  # xlDownThenOver
            
            # Qualité d'impression
            ps.Draft = False
            ps.BlackAndWhite = False
            
        except Exception as e:
            logger.warning(f"⚠️ Configuration mise en page partielle: {e}")
    
    def _get_export_params(self, options: Dict) -> Dict:
        """Obtient les paramètres d'export selon les options"""
        quality_map = {
            'minimum': 1,  # xlQualityMinimum
            'standard': 0,  # xlQualityStandard
            'maximum': 2    # xlQualityMaximum (non disponible partout)
        }
        
        return {
            'quality': quality_map.get(options['quality'], 0)
        }
    
    def _format_size(self, size_bytes: int) -> str:
        """Formate la taille en format lisible"""
        for unit in ['B', 'KB', 'MB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} GB"
    
    def batch_convert(self, excel_files: list, output_dir: str = None) -> list:
        """
        Convertit plusieurs fichiers Excel en PDF
        
        Args:
            excel_files: Liste des chemins Excel
            output_dir: Dossier de sortie (optionnel)
        
        Returns:
            Liste des résultats de conversion
        """
        results = []
        
        if output_dir:
            Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        logger.info(f"🔄 Conversion batch de {len(excel_files)} fichiers")
        
        for i, excel_file in enumerate(excel_files, 1):
            logger.info(f"\n[{i}/{len(excel_files)}] Traitement de {Path(excel_file).name}")
            
            pdf_path = None
            if output_dir:
                pdf_path = Path(output_dir) / f"{Path(excel_file).stem}.pdf"
            
            result = self.convert_excel_to_pdf(excel_file, str(pdf_path) if pdf_path else None)
            results.append(result)
            
            # Pause entre conversions pour éviter surcharge
            if i < len(excel_files):
                time.sleep(1)
        
        # Résumé
        successful = sum(1 for r in results if r['success'])
        logger.info(f"\n📊 Résumé: {successful}/{len(excel_files)} conversions réussies")
        
        return results
    
    def get_stats(self) -> Dict:
        """Retourne les statistiques de conversion"""
        return {
            'total_conversions': self.conversion_stats['total'],
            'successful': self.conversion_stats['success'],
            'failed': self.conversion_stats['failed'],
            'success_rate': (
                self.conversion_stats['success'] / self.conversion_stats['total'] * 100
                if self.conversion_stats['total'] > 0 else 0
            )
        }
