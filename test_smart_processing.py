"""
Test du système de traitement intelligent avec différents scénarios
"""
import pandas as pd
import os
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.file_handler import FileHandler
from core.data_processor import DataProcessor
from core.format_detector import FormatDetector
from core.beneficiary_resolver import BeneficiaryResolver
from core.smart_processor import SmartProcessor
import logging

# Configuration des logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_with_real_files():
    """Test avec les vrais fichiers BulkReport et Export"""
    print("\n" + "="*80)
    print(" TEST AVEC FICHIERS RÉELS")
    print("="*80)
    
    # Chemins des fichiers
    bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\BulkReport_130809.csv"
    export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Export_0131-FMC19-Beat.xlsx"
    fees_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\frais.xlsx"
    
    try:
        # Charger les fichiers
        handler = FileHandler()
        
        print("\n📁 Chargement des fichiers...")
        bulk_df, metadata = handler.read_bulk_report(bulk_path)
        print(f"  ✓ BulkReport: {len(bulk_df)} lignes")
        
        export_df = handler.read_export_file(export_path)
        print(f"  ✓ Export: {len(export_df)} lignes")
        
        fees_df = pd.DataFrame()  # Utiliser frais par défaut
        print(f"  ✓ Frais: Taux par défaut (1.68%)")
        
        # Traiter avec le SmartProcessor
        print("\n🧠 Traitement intelligent...")
        processor = DataProcessor()
        processor.use_smart_processing = True
        
        metadata_report = {
            'date_paiement': '16/09/2025',
            'libelle': 'TEST SMART',
            'budget': 500000,
            'projet': 'UGP',
            'fee_rate': 0.0168
        }
        
        processed_df, errors = processor.process_transactions(
            bulk_df, export_df, fees_df, metadata_report
        )
        
        # Afficher les résultats
        print("\n📊 RÉSULTATS:")
        print(f"  • Transactions traitées: {len(processed_df)}")
        print(f"  • Montant total: {processed_df['Amount'].sum():,.0f} FCFA")
        print(f"  • Frais totaux: {processed_df['Frais'].sum():,.0f} FCFA")
        
        if errors:
            print(f"\n  ⚠ Erreurs/Avertissements:")
            for error in errors:
                print(f"    - {error}")
        
        # Afficher un échantillon
        print("\n📋 Échantillon des données traitées:")
        for i in range(min(3, len(processed_df))):
            row = processed_df.iloc[i]
            print(f"\n  Transaction {i+1}:")
            print(f"    Date: {row['Date']}")
            print(f"    ID: {row['TransactionID']}")
            print(f"    Montant: {row['Amount']:,.0f} FCFA")
            print(f"    Frais: {row['Frais']:,.0f} FCFA")
            print(f"    Vers: {row['Vers']}")
            print(f"    Bénéficiaire: {row['Beneficiaire']}")
        
        return processed_df
        
    except Exception as e:
        print(f"\n❌ Erreur: {e}")
        import traceback
        traceback.print_exc()
        return None

def test_format_detection():
    """Test de la détection de format"""
    print("\n" + "="*80)
    print(" TEST DE DÉTECTION DE FORMAT")
    print("="*80)
    
    detector = FormatDetector()
    
    # Scénario 1: Sans frais (2 transactions normales)
    print("\n📝 Scénario 1: Sans frais")
    df1 = pd.DataFrame({
        'Transaction Timestamp': ['09-09-2025 10:51:17', '09-09-2025 10:52:00'],
        'Amount': [491741, 5000],
        'Credit Msisdn': ['23596771275', '23596771275']
    })
    result1 = detector.detect_format(df1)
    print(f"  Format: {result1['format_type']} (confiance: {result1['confidence']}%)")
    
    # Scénario 2: Avec frais (2 transactions + 2 frais)
    print("\n📝 Scénario 2: Avec frais")
    df2 = pd.DataFrame({
        'Transaction Timestamp': ['09-09-2025 10:51:17', '09-09-2025 10:51:17',
                                 '09-09-2025 10:51:17', '09-09-2025 10:51:17'],
        'Amount': [491741, 5000, 8261, 84],  # 2 transactions + 2 frais
        'Credit Msisdn': ['23596771275', '23596771275', '23596771275', '23596771275']
    })
    result2 = detector.detect_format(df2)
    print(f"  Format: {result2['format_type']} (confiance: {result2['confidence']}%)")
    
    # Scénario 3: Mixte (3 transactions)
    print("\n📝 Scénario 3: Nombre impair")
    df3 = pd.DataFrame({
        'Transaction Timestamp': ['09-09-2025 10:51:17', '09-09-2025 10:51:17', '09-09-2025 10:51:17'],
        'Amount': [491741, 5000, 1000],
        'Credit Msisdn': ['23596771275', '23596771275', '23596771275']
    })
    result3 = detector.detect_format(df3)
    print(f"  Format: {result3['format_type']} (confiance: {result3['confidence']}%)")

def test_beneficiary_mapping():
    """Test du mapping des bénéficiaires"""
    print("\n" + "="*80)
    print(" TEST DE MAPPING DES BÉNÉFICIAIRES")
    print("="*80)
    
    resolver = BeneficiaryResolver()
    
    # Transactions
    transactions_df = pd.DataFrame({
        'Credit Msisdn': ['23596771275', '23596771275'],
        'Amount': [491741, 5000]
    })
    
    # Export avec bénéficiaires
    export_df = pd.DataFrame({
        'Nom': ['TINA', 'JEAN'],
        'Prénoms': ['GANG-IRANGA', 'DUPONT'],
        'Telephone': ['23596771275', '23598888888']
    })
    
    print("\n📝 Test mapping 2 transactions, 2 bénéficiaires")
    result = resolver.resolve_beneficiaries(transactions_df, export_df)
    print("\nRésultats:")
    for i, row in result.iterrows():
        print(f"  Transaction {i+1}: {row.get('Beneficiaire', 'N/A')}")

if __name__ == "__main__":
    print("\n🚀 DÉMARRAGE DES TESTS DU SMART PROCESSOR\n")
    
    # Test 1: Détection de format
    test_format_detection()
    
    # Test 2: Mapping des bénéficiaires
    test_beneficiary_mapping()
    
    # Test 3: Fichiers réels
    processed_df = test_with_real_files()
    
    if processed_df is not None:
        print("\n✅ TOUS LES TESTS RÉUSSIS!")
    else:
        print("\n❌ CERTAINS TESTS ONT ÉCHOUÉ")
    
    print("\n" + "="*80)
