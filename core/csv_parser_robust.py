"""
Parser robuste pour le format spécial du BulkReport
"""
import pandas as pd
import re
import logging

logger = logging.getLogger(__name__)


def parse_bulkreport_robust(file_path):
    """
    Parse le BulkReport avec son format spécifique
    Format: "	1,""	Success"",""	23596771275"",...
    """
    logger.info(f"Parsing robuste de: {file_path}")
    
    # Lire tout le fichier
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Headers à la ligne 13 (index 12)
    headers = ['Record No', 'Validation Result', 'Credit Msisdn', 'Transaction Timestamp', 
               'Finished Timestamp', 'TransactionID', 'Transaction Details', 'Amount', 
               'Fee Charge', 'Extra Fee Charge', 'Tax', 'Status', 'Error Code', 'Error Message']
    
    # Parser les lignes de données (14+)
    data = []
    for i in range(13, len(lines)):
        line = lines[i].strip()
        if not line or line == '""':
            continue
            
        logger.info(f"Parsing ligne {i+1}: {line[:50]}...")
        
        # Méthode 1: Regex pour extraire les valeurs entre guillemets doubles
        # Pattern: capture tout ce qui est entre "" (en gérant les doubles "")
        pattern = r'"([^"]*(?:""[^"]*)*)"'
        matches = re.findall(pattern, line)
        
        if matches:
            # Nettoyer chaque match
            cleaned = []
            for match in matches:
                # Remplacer les doubles guillemets par simple
                value = match.replace('""', '"')
                # Enlever tabs et espaces
                value = value.strip().strip('\t').strip()
                cleaned.append(value)
            
            logger.info(f"  Trouvé {len(cleaned)} valeurs")
            
            # Si on a assez de colonnes
            if len(cleaned) >= 12:
                data.append(cleaned[:14])  # Prendre les 14 premières colonnes
                logger.info(f"  ✓ Transaction: {cleaned[0]} - {cleaned[5]} - {cleaned[7]}")
        else:
            # Méthode 2: Simple split si pas de matches regex
            logger.info(f"  Tentative split simple...")
            
            # Enlever le premier et dernier guillemet
            if line.startswith('"') and line.endswith('"'):
                line = line[1:-1]
            
            # Splitter par ,""
            parts = line.split(',""')
            
            cleaned = []
            for part in parts:
                # Nettoyer
                value = part.strip('"').strip('\t').strip()
                cleaned.append(value)
            
            if len(cleaned) >= 12:
                data.append(cleaned[:14])
                logger.info(f"  ✓ Transaction (split): {cleaned[0]} - Amount: {cleaned[7] if len(cleaned) > 7 else 'N/A'}")
    
    # Créer le DataFrame
    if data:
        df = pd.DataFrame(data, columns=headers)
        logger.info(f"✅ Créé DataFrame avec {len(df)} transactions")
        
        # Nettoyer les colonnes numériques
        numeric_cols = ['Amount', 'Fee Charge', 'Extra Fee Charge', 'Tax']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        return df
    else:
        logger.warning("❌ Aucune donnée parsée")
        return pd.DataFrame(columns=headers)


def test_parser():
    """Test le parser avec le vrai fichier"""
    file_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\BulkReport_130809.csv"
    
    df = parse_bulkreport_robust(file_path)
    
    print(f"\nRésultat du parsing:")
    print(f"Nombre de lignes: {len(df)}")
    
    if len(df) > 0:
        print(f"\nPremières transactions:")
        for i in range(min(3, len(df))):
            print(f"\n  Transaction {i+1}:")
            print(f"    ID: {df.iloc[i]['TransactionID']}")
            print(f"    Montant: {df.iloc[i]['Amount']}")
            print(f"    Téléphone: {df.iloc[i]['Credit Msisdn']}")
            print(f"    Status: {df.iloc[i]['Status']}")
    
    return df


if __name__ == "__main__":
    # Tester directement
    import logging
    logging.basicConfig(level=logging.INFO)
    test_parser()
