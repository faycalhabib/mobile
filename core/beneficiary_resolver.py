"""
Module de résolution intelligente des bénéficiaires
Gère le mapping entre transactions et bénéficiaires selon différentes stratégies
"""
import pandas as pd
import logging
from typing import List, Dict, Tuple

logger = logging.getLogger(__name__)


class BeneficiaryResolver:
    """Résout le mapping entre transactions et bénéficiaires"""
    
    def __init__(self):
        self.mapping_strategy = None
        self.mapping_stats = {}
    
    def resolve_beneficiaries(self, transactions_df: pd.DataFrame, 
                             export_df: pd.DataFrame) -> pd.DataFrame:
        """
        Mappe intelligemment les bénéficiaires aux transactions
        
        Returns:
            DataFrame avec colonne 'Beneficiaire' ajoutée/mise à jour
        """
        logger.info("="*60)
        logger.info("RÉSOLUTION DES BÉNÉFICIAIRES")
        logger.info(f"Transactions: {len(transactions_df)} | Bénéficiaires disponibles: {len(export_df)}")
        
        # Analyser la situation
        strategy = self._determine_strategy(transactions_df, export_df)
        self.mapping_strategy = strategy['name']
        
        logger.info(f"Stratégie sélectionnée: {strategy['name']}")
        logger.info(f"Raison: {strategy['reason']}")
        
        # Appliquer la stratégie
        if strategy['name'] == 'ONE_TO_ONE':
            result = self._map_one_to_one(transactions_df, export_df)
        elif strategy['name'] == 'BY_PHONE_NUMBER':
            result = self._map_by_phone(transactions_df, export_df)
        elif strategy['name'] == 'WITH_DUPLICATION':
            result = self._map_with_duplication(transactions_df, export_df)
        elif strategy['name'] == 'PARTIAL':
            result = self._map_partial(transactions_df, export_df)
        else:
            result = self._map_fallback(transactions_df)
        
        # Statistiques de mapping
        self._log_mapping_stats(result)
        
        logger.info("="*60)
        return result
    
    def _determine_strategy(self, transactions_df: pd.DataFrame, 
                           export_df: pd.DataFrame) -> Dict:
        """Détermine la meilleure stratégie de mapping"""
        
        t_count = len(transactions_df)
        e_count = len(export_df)
        
        # Identifier la colonne des numéros de téléphone
        phone_col = self._find_phone_column(transactions_df)
        
        if t_count == 0:
            return {
                'name': 'NONE',
                'reason': 'Aucune transaction à mapper'
            }
        
        if e_count == 0:
            return {
                'name': 'FALLBACK',
                'reason': 'Aucun bénéficiaire dans Export'
            }
        
        # Cas 1: Même nombre de lignes → mapping 1-pour-1
        if t_count == e_count:
            return {
                'name': 'ONE_TO_ONE',
                'reason': f'Nombre égal ({t_count} transactions = {e_count} bénéficiaires)'
            }
        
        # Cas 2: Plus de bénéficiaires que de transactions
        if e_count > t_count:
            return {
                'name': 'BY_PHONE_NUMBER',
                'reason': f'Plus de bénéficiaires ({e_count}) que de transactions ({t_count})'
            }
        
        # Cas 3: Plus de transactions que de bénéficiaires
        if t_count > e_count and e_count > 0:
            # Vérifier si on peut mapper par numéro
            if phone_col:
                unique_phones = transactions_df[phone_col].nunique()
                if unique_phones <= e_count:
                    return {
                        'name': 'WITH_DUPLICATION',
                        'reason': f'{t_count} transactions vers {unique_phones} numéros uniques'
                    }
            
            return {
                'name': 'PARTIAL',
                'reason': f'Seulement {e_count} bénéficiaires pour {t_count} transactions'
            }
        
        return {
            'name': 'FALLBACK',
            'reason': 'Cas non géré, utilisation du fallback'
        }
    
    def _map_one_to_one(self, transactions_df: pd.DataFrame, 
                       export_df: pd.DataFrame) -> pd.DataFrame:
        """Mapping direct 1-pour-1 par index"""
        logger.info("  → Mapping 1-pour-1 par ordre d'apparition")
        
        result = transactions_df.copy()
        
        # Extraire directement les noms depuis la colonne "Nom et prénoms" ou "Nom"
        names = []
        
        # Chercher la bonne colonne dans Export
        name_column = None
        for col in export_df.columns:
            col_lower = col.lower()
            if 'nom' in col_lower and 'prénom' in col_lower:
                name_column = col
                break
            elif 'nom' in col_lower:
                name_column = col
                break
        
        # Si pas trouvé, prendre la première colonne texte
        if not name_column:
            for col in export_df.columns:
                if export_df[col].dtype == 'object':
                    name_column = col
                    break
        
        if name_column:
            logger.info(f"  → Utilisation de la colonne: {name_column}")
            for idx, row in export_df.iterrows():
                if pd.notna(row[name_column]):
                    names.append(str(row[name_column]).strip())
        
        logger.info(f"  → {len(names)} noms extraits depuis Export")
        
        # Mapper chaque transaction avec un nom
        mapped = 0
        for i in range(len(result)):
            if i < len(names):
                result.loc[result.index[i], 'Beneficiaire'] = names[i]
                logger.info(f"    Transaction {i+1} → {names[i]}")
                mapped += 1
            else:
                # Si plus de transactions que de bénéficiaires, recycler les noms
                if names:
                    recycled_index = i % len(names)
                    result.loc[result.index[i], 'Beneficiaire'] = names[recycled_index]
                    logger.info(f"    Transaction {i+1} → {names[recycled_index]} (recyclé)")
                    mapped += 1
                else:
                    result.loc[result.index[i], 'Beneficiaire'] = f"BENEFICIAIRE_{i+1}"
                    logger.warning(f"    Transaction {i+1} → Pas de bénéficiaire")
        
        self.mapping_stats = {
            'mapped': mapped,
            'unmapped': len(result) - mapped,
            'method': 'one-to-one'
        }
        
        return result
    
    def _map_by_phone(self, transactions_df: pd.DataFrame, 
                     export_df: pd.DataFrame) -> pd.DataFrame:
        """Mapping par numéro de téléphone"""
        logger.info("  → Mapping par numéro de téléphone")
        
        result = transactions_df.copy()
        phone_col = self._find_phone_column(result)
        
        if not phone_col:
            logger.warning("  ⚠ Pas de colonne téléphone trouvée")
            return self._map_fallback(result)
        
        # Créer le dictionnaire de mapping
        phone_map = self._create_phone_map(export_df)
        
        # Mapper les bénéficiaires
        mapped = 0
        unmapped = 0
        
        for idx, row in result.iterrows():
            phone = str(row[phone_col]).strip()
            if phone in phone_map:
                result.loc[idx, 'Beneficiaire'] = phone_map[phone]
                logger.info(f"    {phone} → {phone_map[phone]}")
                mapped += 1
            else:
                result.loc[idx, 'Beneficiaire'] = f"BENEFICIAIRE_{phone[-4:]}"
                logger.warning(f"    {phone} → Non trouvé, utilise BENEFICIAIRE_{phone[-4:]}")
                unmapped += 1
        
        self.mapping_stats = {
            'mapped': mapped,
            'unmapped': unmapped,
            'method': 'by-phone'
        }
        
        return result
    
    def _map_with_duplication(self, transactions_df: pd.DataFrame, 
                             export_df: pd.DataFrame) -> pd.DataFrame:
        """Mapping avec duplication si plusieurs transactions vers même numéro"""
        logger.info("  → Mapping avec duplication pour transactions multiples")
        
        result = transactions_df.copy()
        phone_col = self._find_phone_column(result)
        
        if not phone_col:
            return self._map_one_to_one(transactions_df, export_df)
        
        # Créer mapping par numéro
        phone_map = self._create_phone_map(export_df)
        
        # Grouper transactions par numéro
        phone_groups = result.groupby(phone_col).groups
        
        mapped = 0
        for phone, indices in phone_groups.items():
            phone_str = str(phone).strip()
            if phone_str in phone_map:
                name = phone_map[phone_str]
                for idx in indices:
                    result.loc[idx, 'Beneficiaire'] = name
                    mapped += 1
                logger.info(f"    {phone_str} ({len(indices)} transactions) → {name}")
            else:
                for i, idx in enumerate(indices):
                    result.loc[idx, 'Beneficiaire'] = f"BENEFICIAIRE_{phone_str[-4:]}_{i+1}"
                logger.warning(f"    {phone_str} ({len(indices)} transactions) → Non trouvé")
        
        self.mapping_stats = {
            'mapped': mapped,
            'unmapped': len(result) - mapped,
            'method': 'with-duplication'
        }
        
        return result
    
    def _map_partial(self, transactions_df: pd.DataFrame, 
                    export_df: pd.DataFrame) -> pd.DataFrame:
        """Mapping partiel quand pas assez de bénéficiaires"""
        logger.info("  → Mapping partiel (pas assez de bénéficiaires)")
        
        # Utiliser le mapping one-to-one avec recyclage
        return self._map_one_to_one(transactions_df, export_df)
    
    def _map_fallback(self, transactions_df: pd.DataFrame) -> pd.DataFrame:
        """Mapping de secours quand aucun bénéficiaire disponible"""
        logger.warning("  → Utilisation du mapping de secours")
        
        result = transactions_df.copy()
        
        for i, idx in enumerate(result.index):
            result.loc[idx, 'Beneficiaire'] = f"BENEFICIAIRE_{i+1}"
        
        self.mapping_stats = {
            'mapped': 0,
            'unmapped': len(result),
            'method': 'fallback'
        }
        
        return result
    
    def _find_phone_column(self, df: pd.DataFrame) -> str:
        """Trouve la colonne contenant les numéros de téléphone"""
        possible_cols = ['Credit Msisdn', 'Debit Msisdn', 'Receiver', 
                        'Destination', 'Vers', 'To', 'Phone', 'Telephone']
        
        for col in possible_cols:
            if col in df.columns:
                return col
        
        # Chercher par pattern
        for col in df.columns:
            if 'phone' in col.lower() or 'msisdn' in col.lower() or 'tel' in col.lower():
                return col
        
        return None
    
    def _extract_names(self, export_df: pd.DataFrame) -> List[str]:
        """Extrait les noms depuis le DataFrame Export"""
        names = []
        
        # Chercher les colonnes de noms
        name_cols = []
        for col in export_df.columns:
            col_lower = col.lower()
            if any(word in col_lower for word in ['nom', 'name', 'beneficiaire', 'prenoms', 'firstname', 'lastname']):
                name_cols.append(col)
        
        if not name_cols:
            # Prendre la première colonne non numérique
            for col in export_df.columns:
                if export_df[col].dtype == 'object':
                    name_cols.append(col)
                    break
        
        if name_cols:
            # Concaténer les colonnes de noms
            for idx, row in export_df.iterrows():
                name_parts = []
                for col in name_cols:
                    if pd.notna(row[col]):
                        name_parts.append(str(row[col]).strip())
                
                if name_parts:
                    names.append(' '.join(name_parts))
                else:
                    names.append(f"BENEFICIAIRE_{idx+1}")
        
        return names
    
    def _create_phone_map(self, export_df: pd.DataFrame) -> Dict[str, str]:
        """Crée un dictionnaire téléphone -> nom"""
        phone_map = {}
        
        # Trouver la colonne téléphone dans Export
        phone_col = self._find_phone_column(export_df)
        names = self._extract_names(export_df)
        
        if phone_col and phone_col in export_df.columns:
            for i, (idx, row) in enumerate(export_df.iterrows()):
                if i < len(names) and pd.notna(row[phone_col]):
                    phone = str(row[phone_col]).strip()
                    phone_map[phone] = names[i]
        else:
            # Si pas de colonne téléphone, utiliser l'index
            logger.warning("  ⚠ Pas de colonne téléphone dans Export")
        
        return phone_map
    
    def _log_mapping_stats(self, result_df: pd.DataFrame):
        """Log les statistiques de mapping"""
        logger.info("\n📊 STATISTIQUES DE MAPPING:")
        logger.info(f"  • Stratégie: {self.mapping_strategy}")
        logger.info(f"  • Transactions mappées: {self.mapping_stats.get('mapped', 0)}")
        logger.info(f"  • Transactions non mappées: {self.mapping_stats.get('unmapped', 0)}")
        
        # Vérifier les valeurs manquantes
        if 'Beneficiaire' in result_df.columns:
            missing = result_df['Beneficiaire'].isna().sum()
            placeholders = result_df['Beneficiaire'].str.startswith('BENEFICIAIRE_').sum()
            logger.info(f"  • Valeurs manquantes: {missing}")
            logger.info(f"  • Placeholders utilisés: {placeholders}")
