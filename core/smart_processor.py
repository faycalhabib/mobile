"""
Processeur intelligent qui orchestre la détection de format et le mapping des bénéficiaires
"""
import pandas as pd
import logging
from typing import Tuple, Dict, Any
from .format_detector import FormatDetector
from .beneficiary_resolver_v2 import BeneficiaryResolverV2 as BeneficiaryResolver

logger = logging.getLogger(__name__)


class SmartProcessor:
    """Orchestre le traitement intelligent des données"""
    
    def __init__(self):
        self.format_detector = FormatDetector()
        self.beneficiary_resolver = BeneficiaryResolver()
        self.processing_stats = {}
    
    def process_smart(self, 
                     bulk_df: pd.DataFrame,
                     export_df: pd.DataFrame,
                     fees_df: pd.DataFrame,
                     metadata: dict) -> Tuple[pd.DataFrame, Dict[str, Any]]:
        """
        Traite intelligemment les données avec détection automatique et mapping adaptatif
        
        Returns:
            Tuple: (DataFrame traité, statistiques de traitement)
        """
        logger.info("\n" + "="*70)
        logger.info(" SMART PROCESSOR - TRAITEMENT INTELLIGENT DES DONNÉES")
        logger.info("="*70)
        
        stats = {
            'original_count': len(bulk_df),
            'format_detected': None,
            'transactions_kept': 0,
            'fees_filtered': 0,
            'beneficiaries_mapped': 0,
            'beneficiaries_missing': 0,
            'errors': []
        }
        
        try:
            # Étape 1: Détection du format
            logger.info("\n📋 ÉTAPE 1: DÉTECTION DU FORMAT")
            format_info = self.format_detector.detect_format(bulk_df)
            stats['format_detected'] = format_info['format_type']
            stats['confidence'] = format_info['confidence']
            
            # Étape 2: Filtrage selon le format
            logger.info("\n🔄 ÉTAPE 2: FILTRAGE DES TRANSACTIONS")
            filtered_df = self.format_detector.apply_filter(bulk_df, format_info)
            stats['transactions_kept'] = len(filtered_df)
            stats['fees_filtered'] = len(bulk_df) - len(filtered_df)
            
            logger.info(f"  • Transactions gardées: {stats['transactions_kept']}")
            logger.info(f"  • Lignes de frais filtrées: {stats['fees_filtered']}")
            
            # Étape 3: Mapping des bénéficiaires
            logger.info(f"\n👥 ÉTAPE 3: MAPPING DES BÉNÉFICIAIRES")
            result = self.beneficiary_resolver.resolve_beneficiaries(filtered_df, export_df)
            
            # Ajouter les colonnes mappées au DataFrame filtré
            for col in result.columns:
                if col not in filtered_df.columns:
                    filtered_df[col] = result[col]
            
            # S'assurer que les colonnes essentielles sont présentes
            # Date - extraire de Transaction Timestamp ou Finished Timestamp
            if 'Transaction Timestamp' in filtered_df.columns:
                filtered_df['Date'] = filtered_df['Transaction Timestamp']
            elif 'Finished Timestamp' in filtered_df.columns:
                filtered_df['Date'] = filtered_df['Finished Timestamp']
            
            # Vers - utiliser Credit Msisdn
            if 'Credit Msisdn' in filtered_df.columns:
                filtered_df['Vers'] = filtered_df['Credit Msisdn']
            
            # Étape 4: Calcul des frais
            logger.info("\n💰 ÉTAPE 4: CALCUL DES FRAIS")
            processed_df = self._calculate_fees(filtered_df, fees_df, metadata)
            
            # Étape 5: Validation finale
            logger.info("\n✅ ÉTAPE 5: VALIDATION FINALE")
            validation_results = self._validate_output(processed_df)
            stats.update(validation_results)
            
            # Statistiques finales
            self._log_final_stats(processed_df, stats)
            
            self.processing_stats = stats
            return processed_df, stats
            
        except Exception as e:
            logger.error(f"❌ Erreur dans le traitement intelligent: {str(e)}")
            stats['errors'].append(str(e))
            raise
    
    def _calculate_fees(self, df: pd.DataFrame, fees_df: pd.DataFrame, metadata: dict) -> pd.DataFrame:
        """Calcule les frais pour chaque transaction"""
        
        # Taux par défaut
        default_rate = metadata.get('fee_rate', 0.0168)  # 1.68%
        
        if fees_df is not None and not fees_df.empty:
            # Utiliser la table des frais
            logger.info("  → Utilisation de la table des frais")
            df['Frais'] = df.apply(lambda row: self._get_fee_from_table(row, fees_df, default_rate), axis=1)
        else:
            # Utiliser le taux par défaut
            logger.info(f"  → Utilisation du taux par défaut ({default_rate*100:.2f}%)")
            df['Frais'] = df['Amount'] * default_rate
        
        # Arrondir les frais
        df['Frais'] = df['Frais'].round(0).astype(int)
        
        total_fees = df['Frais'].sum()
        avg_fee_rate = (df['Frais'].sum() / df['Amount'].sum() * 100) if df['Amount'].sum() > 0 else 0
        
        logger.info(f"  • Total des frais calculés: {total_fees:,.0f} FCFA")
        logger.info(f"  • Taux moyen appliqué: {avg_fee_rate:.2f}%")
        
        return df
    
    def _get_fee_from_table(self, row, fees_df: pd.DataFrame, default_rate: float) -> float:
        """Obtient les frais depuis la table ou utilise le taux par défaut"""
        amount = row.get('Amount', 0)
        
        # Chercher dans la table des frais par tranche
        for _, fee_row in fees_df.iterrows():
            if 'min_amount' in fee_row and 'max_amount' in fee_row:
                if fee_row['min_amount'] <= amount <= fee_row['max_amount']:
                    if 'fee_amount' in fee_row:
                        return fee_row['fee_amount']
                    elif 'fee_rate' in fee_row:
                        return amount * fee_row['fee_rate']
        
        # Si pas trouvé, utiliser le taux par défaut
        return amount * default_rate
    
    def _validate_output(self, df: pd.DataFrame) -> Dict:
        """Valide le DataFrame de sortie"""
        results = {
            'is_valid': True,
            'warnings': [],
            'missing_columns': []
        }
        
        # Colonnes requises
        required_columns = ['Date', 'TransactionID', 'Amount', 'Vers', 'Beneficiaire', 'Frais']
        
        for col in required_columns:
            if col not in df.columns:
                results['missing_columns'].append(col)
                results['is_valid'] = False
                logger.warning(f"  ⚠ Colonne manquante: {col}")
        
        # Vérifications des données
        if 'Amount' in df.columns:
            if (df['Amount'] <= 0).any():
                results['warnings'].append("Montants négatifs ou zéro détectés")
                logger.warning("  ⚠ Montants négatifs ou zéro détectés")
        
        if 'Beneficiaire' in df.columns:
            missing = df['Beneficiaire'].isna().sum()
            if missing > 0:
                results['warnings'].append(f"{missing} bénéficiaires manquants")
                logger.warning(f"  ⚠ {missing} bénéficiaires manquants")
            
            placeholders = df['Beneficiaire'].str.startswith('BENEFICIAIRE_', na=False).sum()
            if placeholders > 0:
                results['warnings'].append(f"{placeholders} placeholders utilisés")
                logger.info(f"  ℹ {placeholders} placeholders utilisés")
        
        if results['is_valid']:
            logger.info("  ✓ Validation réussie")
        else:
            logger.error("  ✗ Validation échouée")
        
        return results
    
    def _log_final_stats(self, df: pd.DataFrame, stats: Dict):
        """Log les statistiques finales"""
        logger.info("\n" + "="*70)
        logger.info(" 📊 STATISTIQUES FINALES")
        logger.info("="*70)
        
        logger.info(f"  Format détecté: {stats['format_detected']} (confiance: {stats.get('confidence', 0)}%)")
        logger.info(f"  Transactions originales: {stats['original_count']}")
        logger.info(f"  Transactions gardées: {stats['transactions_kept']}")
        logger.info(f"  Lignes de frais filtrées: {stats['fees_filtered']}")
        
        if 'Amount' in df.columns:
            total_amount = df['Amount'].sum()
            logger.info(f"  Montant total: {total_amount:,.0f} FCFA")
        
        if 'Frais' in df.columns:
            total_fees = df['Frais'].sum()
            logger.info(f"  Frais totaux: {total_fees:,.0f} FCFA")
        
        if 'Beneficiaire' in df.columns:
            unique_beneficiaries = df['Beneficiaire'].nunique()
            logger.info(f"  Bénéficiaires uniques: {unique_beneficiaries}")
        
        if stats.get('warnings'):
            logger.info(f"\n  ⚠ Avertissements:")
            for warning in stats['warnings']:
                logger.info(f"    - {warning}")
        
        logger.info("="*70)
