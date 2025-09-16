"""
Module amélioré pour la résolution des bénéficiaires avec logs détaillés
"""
import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


class BeneficiaryResolverV2:
    """Résolution intelligente des bénéficiaires version 2 avec logs détaillés"""
    
    def __init__(self):
        self.mapping_strategy = None
        self.mapping_stats = {}
        
    def resolve_beneficiaries(self, transactions_df: pd.DataFrame, 
                             export_df: pd.DataFrame) -> pd.DataFrame:
        """
        Résout les bénéficiaires pour chaque transaction avec logs détaillés
        """
        logger.info("="*60)
        logger.info("DÉBUT DE LA RÉSOLUTION DES BÉNÉFICIAIRES V2")
        logger.info("="*60)
        
        # 1. Analyser les DataFrames d'entrée
        logger.info(f"📊 Données d'entrée:")
        logger.info(f"  • Transactions: {len(transactions_df)} lignes")
        logger.info(f"  • Export: {len(export_df)} lignes")
        
        # 2. Chercher et extraire les noms depuis Export
        names = self._extract_names_robust(export_df)
        
        if not names:
            logger.error("❌ AUCUN NOM EXTRAIT depuis Export!")
            return self._apply_fallback(transactions_df)
        
        logger.info(f"✅ {len(names)} noms extraits avec succès")
        for i, name in enumerate(names[:3]):
            logger.info(f"    {i+1}. {name}")
        if len(names) > 3:
            logger.info(f"    ... et {len(names) - 3} autres")
        
        # 3. Appliquer le mapping
        result = self._apply_mapping(transactions_df, names)
        
        # 4. Statistiques finales
        self._log_final_stats(result)
        
        return result
    
    def _extract_names_robust(self, export_df: pd.DataFrame) -> List[str]:
        """
        Extraction robuste des noms avec plusieurs stratégies
        """
        logger.info("\n🔍 EXTRACTION DES NOMS DEPUIS EXPORT")
        logger.info("-"*40)
        
        names = []
        
        # Afficher toutes les colonnes disponibles
        logger.info("Colonnes disponibles dans Export:")
        for col in export_df.columns:
            logger.info(f"  • '{col}' (type: {export_df[col].dtype})")
        
        # STRATÉGIE 1: Chercher une colonne avec 'nom' et 'prénom'
        name_column = None
        for col in export_df.columns:
            col_lower = col.lower()
            # Vérifier différentes variantes
            if ('nom' in col_lower and 'prénom' in col_lower) or \
               ('nom' in col_lower and 'prenom' in col_lower):
                name_column = col
                logger.info(f"✅ Stratégie 1: Colonne trouvée '{col}'")
                break
        
        # STRATÉGIE 2: Chercher juste 'nom'
        if not name_column:
            for col in export_df.columns:
                if 'nom' in col.lower():
                    name_column = col
                    logger.info(f"✅ Stratégie 2: Colonne trouvée '{col}'")
                    break
        
        # STRATÉGIE 3: Première colonne non numérique
        if not name_column:
            for col in export_df.columns:
                if export_df[col].dtype == 'object' or export_df[col].dtype == 'string':
                    # Vérifier que ce n'est pas une colonne de téléphone
                    if not any(word in col.lower() for word in ['tel', 'phone', 'msisdn', 'numero']):
                        name_column = col
                        logger.info(f"✅ Stratégie 3: Première colonne texte '{col}'")
                        break
        
        # STRATÉGIE 4: Utiliser l'index 0 ou 1
        if not name_column and len(export_df.columns) > 0:
            # Essayer la première colonne
            name_column = export_df.columns[0]
            logger.warning(f"⚠️ Stratégie 4: Utilisation forcée de la première colonne '{name_column}'")
        
        # Extraire les valeurs
        if name_column:
            logger.info(f"\n📋 Extraction depuis la colonne: '{name_column}'")
            
            # Afficher les premières valeurs pour debug
            logger.info("Premières valeurs de cette colonne:")
            for i in range(min(3, len(export_df))):
                val = export_df.iloc[i][name_column]
                logger.info(f"  Ligne {i}: '{val}' (type: {type(val)})")
            
            # Extraire tous les noms
            for idx, row in export_df.iterrows():
                value = row[name_column]
                
                # Vérifier différents cas
                if pd.isna(value):
                    logger.debug(f"  Ligne {idx}: valeur NaN ignorée")
                    continue
                    
                value_str = str(value).strip()
                
                if value_str == '' or value_str == 'nan':
                    logger.debug(f"  Ligne {idx}: valeur vide ignorée")
                    continue
                
                # Nettoyer la valeur
                value_clean = value_str.replace('\n', ' ').replace('\t', ' ').strip()
                
                if value_clean:
                    names.append(value_clean)
                    logger.debug(f"  Ligne {idx}: '{value_clean}' ajouté")
        else:
            logger.error("❌ AUCUNE COLONNE DE NOMS TROUVÉE!")
            
            # Debug: afficher tout le DataFrame
            logger.error("Contenu complet du DataFrame Export:")
            logger.error(export_df.head().to_string())
        
        return names
    
    def _apply_mapping(self, transactions_df: pd.DataFrame, names: List[str]) -> pd.DataFrame:
        """
        Applique le mapping des noms aux transactions
        """
        logger.info("\n🔄 APPLICATION DU MAPPING")
        logger.info("-"*40)
        
        result = transactions_df.copy()
        num_trans = len(result)
        num_names = len(names)
        
        logger.info(f"Mapping de {num_trans} transactions avec {num_names} noms")
        
        # Créer la colonne Beneficiaire
        result['Beneficiaire'] = ''
        
        # Mapper les transactions
        for i in range(num_trans):
            if i < num_names:
                # Mapping direct
                beneficiaire = names[i]
                result.loc[result.index[i], 'Beneficiaire'] = beneficiaire
                logger.info(f"  Transaction {i+1} → '{beneficiaire}'")
            else:
                # Recyclage des noms si plus de transactions
                recycled_idx = i % num_names
                beneficiaire = names[recycled_idx]
                result.loc[result.index[i], 'Beneficiaire'] = beneficiaire
                logger.info(f"  Transaction {i+1} → '{beneficiaire}' (recyclé depuis position {recycled_idx+1})")
        
        self.mapping_stats = {
            'total_transactions': num_trans,
            'total_names': num_names,
            'mapped': num_trans,
            'recycled': max(0, num_trans - num_names)
        }
        
        return result
    
    def _apply_fallback(self, transactions_df: pd.DataFrame) -> pd.DataFrame:
        """
        Applique un mapping de secours
        """
        logger.warning("\n⚠️ APPLICATION DU MAPPING DE SECOURS")
        
        result = transactions_df.copy()
        
        # Essayer d'extraire le numéro de téléphone pour un meilleur placeholder
        phone_col = None
        for col in result.columns:
            if 'msisdn' in col.lower() or 'phone' in col.lower():
                phone_col = col
                break
        
        for i in range(len(result)):
            if phone_col and phone_col in result.columns:
                phone = str(result.iloc[i][phone_col])
                # Prendre les 4 derniers chiffres
                last_digits = phone[-4:] if len(phone) >= 4 else phone
                result.loc[result.index[i], 'Beneficiaire'] = f"BENEFICIAIRE_{last_digits}_{i+1}"
            else:
                result.loc[result.index[i], 'Beneficiaire'] = f"BENEFICIAIRE_{i+1}"
        
        return result
    
    def _log_final_stats(self, result_df: pd.DataFrame):
        """
        Affiche les statistiques finales
        """
        logger.info("\n📊 STATISTIQUES FINALES")
        logger.info("-"*40)
        
        if 'Beneficiaire' in result_df.columns:
            # Compter les placeholders
            placeholders = result_df['Beneficiaire'].astype(str).str.contains('BENEFICIAIRE_').sum()
            real_names = len(result_df) - placeholders
            
            logger.info(f"  • Total transactions: {len(result_df)}")
            logger.info(f"  • Vrais noms assignés: {real_names}")
            logger.info(f"  • Placeholders: {placeholders}")
            
            if placeholders > 0:
                logger.warning(f"  ⚠️ {placeholders} transactions ont des placeholders!")
            else:
                logger.info(f"  ✅ Tous les bénéficiaires ont des vrais noms!")
            
            # Afficher un échantillon
            logger.info("\nÉchantillon des bénéficiaires:")
            for i in range(min(5, len(result_df))):
                benef = result_df.iloc[i]['Beneficiaire']
                logger.info(f"  {i+1}. {benef}")
        else:
            logger.error("❌ Colonne 'Beneficiaire' non créée!")
