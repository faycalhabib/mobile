"""
Module de détection intelligente du format BulkReport
Identifie si le fichier contient des lignes de frais séparées ou non
"""
import pandas as pd
import logging
from datetime import datetime

logger = logging.getLogger(__name__)


class FormatDetector:
    """Détecte automatiquement le format du BulkReport"""
    
    def detect_format(self, df: pd.DataFrame) -> dict:
        """
        Analyse le DataFrame pour déterminer son format
        
        Returns:
            dict: {
                'format_type': 'WITH_FEES' ou 'WITHOUT_FEES',
                'transaction_count': nombre de transactions principales,
                'fee_count': nombre de lignes de frais,
                'confidence': niveau de confiance (0-100),
                'details': détails de l'analyse
            }
        """
        logger.info("="*60)
        logger.info("DÉBUT DÉTECTION FORMAT BULKREPORT")
        logger.info(f"Nombre total de lignes: {len(df)}")
        
        result = {
            'format_type': 'WITHOUT_FEES',
            'transaction_count': len(df),
            'fee_count': 0,
            'confidence': 100,
            'details': []
        }
        
        # Si moins de 2 lignes, pas de frais séparés possibles
        if len(df) < 2:
            logger.info("✓ Moins de 2 lignes → Pas de frais séparés")
            result['details'].append("Une seule transaction détectée")
            return result
        
        # Analyser les patterns
        analysis = self._analyze_patterns(df)
        
        # Décider du format basé sur l'analyse
        if analysis['has_duplicate_timestamps'] and analysis['has_fee_pattern']:
            result['format_type'] = 'WITH_FEES'
            result['transaction_count'] = analysis['estimated_transactions']
            result['fee_count'] = analysis['estimated_fees']
            result['confidence'] = analysis['confidence']
            result['details'] = analysis['reasons']
            logger.info(f"✓ Format détecté: WITH_FEES (confiance: {result['confidence']}%)")
        else:
            logger.info(f"✓ Format détecté: WITHOUT_FEES")
            result['details'] = ["Pas de pattern de frais détecté"]
        
        logger.info("="*60)
        return result
    
    def _analyze_patterns(self, df: pd.DataFrame) -> dict:
        """Analyse détaillée des patterns dans les données"""
        analysis = {
            'has_duplicate_timestamps': False,
            'has_fee_pattern': False,
            'estimated_transactions': len(df),
            'estimated_fees': 0,
            'confidence': 0,
            'reasons': []
        }
        
        # 1. Vérifier les timestamps dupliqués
        if 'Transaction Timestamp' in df.columns or 'Finished Timestamp' in df.columns:
            time_col = 'Transaction Timestamp' if 'Transaction Timestamp' in df.columns else 'Finished Timestamp'
            
            # Convertir en datetime et grouper par seconde
            df['timestamp_sec'] = pd.to_datetime(df[time_col], errors='coerce').dt.floor('s')
            timestamp_groups = df.groupby('timestamp_sec').size()
            
            # Si des groupes ont un nombre pair de lignes → possible format WITH_FEES
            even_groups = timestamp_groups[timestamp_groups % 2 == 0]
            if len(even_groups) > 0:
                analysis['has_duplicate_timestamps'] = True
                analysis['reasons'].append(f"{len(even_groups)} groupes avec nombre pair de transactions")
                logger.info(f"  → {len(even_groups)} groupes de timestamps avec nombre pair")
        
        # 2. Analyser les montants pour détecter les frais
        if 'Amount' in df.columns:
            amounts = df['Amount'].astype(float)
            
            # Si nombre pair de lignes
            if len(df) % 2 == 0:
                mid = len(df) // 2
                first_half = amounts.iloc[:mid]
                second_half = amounts.iloc[mid:]
                
                # Calculer les ratios
                ratios = []
                for i in range(mid):
                    if first_half.iloc[i] > 0:
                        ratio = second_half.iloc[i] / first_half.iloc[i]
                        ratios.append(ratio)
                        logger.info(f"  → Ligne {i+1}: {first_half.iloc[i]:.0f} | Ligne {mid+i+1}: {second_half.iloc[i]:.0f} | Ratio: {ratio:.4f}")
                
                # Si tous les ratios < 5%, c'est probablement des frais
                if ratios and all(r < 0.05 for r in ratios):
                    analysis['has_fee_pattern'] = True
                    analysis['estimated_transactions'] = mid
                    analysis['estimated_fees'] = mid
                    analysis['confidence'] = 95
                    analysis['reasons'].append(f"Tous les ratios < 5% ({min(ratios):.2%} - {max(ratios):.2%})")
                    logger.info(f"  ✓ Pattern de frais détecté! Ratios: {min(ratios):.2%} - {max(ratios):.2%}")
                elif ratios and any(r < 0.05 for r in ratios):
                    # Pattern partiel
                    analysis['confidence'] = 60
                    analysis['reasons'].append(f"Pattern partiel: {sum(r < 0.05 for r in ratios)}/{len(ratios)} ratios < 5%")
            
            # 3. Vérifier si les petits montants sont groupés à la fin
            if len(df) >= 4:
                avg_first_quarter = amounts.iloc[:len(df)//4].mean()
                avg_last_quarter = amounts.iloc[-len(df)//4:].mean()
                
                if avg_last_quarter < avg_first_quarter * 0.1:
                    analysis['has_fee_pattern'] = True
                    analysis['confidence'] = max(analysis['confidence'], 80)
                    analysis['reasons'].append(f"Derniers montants < 10% des premiers")
                    logger.info(f"  → Pattern détecté: derniers montants très faibles")
        
        # 4. Vérifier les numéros de destination
        dest_col = None
        for col in ['Credit Msisdn', 'Debit Msisdn', 'Receiver', 'Destination']:
            if col in df.columns:
                dest_col = col
                break
        
        if dest_col:
            destinations = df[dest_col]
            
            # Si même numéro répété dans les deux moitiés
            if len(df) % 2 == 0:
                mid = len(df) // 2
                first_half_dest = set(destinations.iloc[:mid])
                second_half_dest = set(destinations.iloc[mid:])
                
                if first_half_dest == second_half_dest:
                    analysis['has_duplicate_pattern'] = True
                    analysis['confidence'] = min(100, analysis['confidence'] + 20)
                    analysis['reasons'].append("Mêmes numéros dans les deux moitiés")
                    logger.info(f"  → Mêmes destinations dans les deux moitiés")
        
        return analysis
    
    def apply_filter(self, df: pd.DataFrame, format_info: dict) -> pd.DataFrame:
        """
        Applique le filtrage approprié selon le format détecté
        """
        if format_info['format_type'] == 'WITH_FEES':
            # Garder seulement la première moitié (transactions principales)
            transaction_count = format_info['transaction_count']
            logger.info(f"Filtrage: Garde les {transaction_count} premières lignes (transactions)")
            return df.iloc[:transaction_count].copy()
        else:
            # Garder tout
            logger.info(f"Pas de filtrage: Toutes les {len(df)} lignes sont des transactions")
            return df.copy()
