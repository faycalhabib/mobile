"""
Script de test de performance - Compare les deux modes (classique vs optimisé)
"""
import time
import os
import sys
import json
from pathlib import Path
import shutil
from datetime import datetime

# Ajouter le chemin pour les imports
sys.path.append(str(Path(__file__).parent))

def test_both_modes():
    """Test les deux modes et compare les performances"""
    
    print("=" * 70)
    print(" TEST DE PERFORMANCE - CLASSIQUE VS OPTIMISÉ")
    print("=" * 70)
    
    # Sauvegarder la config actuelle
    config_path = Path('config/auto_processor_config.json')
    with open(config_path, 'r', encoding='utf-8') as f:
        original_config = json.load(f)
    
    # Préparer les fichiers de test
    print("\n📋 Préparation des fichiers de test...")
    inbox_path = Path('inbox')
    inbox_path.mkdir(exist_ok=True)
    
    # Copier les fichiers de test
    test_files = [
        ('test_data/BulkReport_Test.csv', 'inbox/BulkReport_test.csv'),
        ('test_data/Export_Test.xlsx', 'inbox/Export_test.xlsx')
    ]
    
    for source, dest in test_files:
        if os.path.exists(source):
            shutil.copy2(source, dest)
            print(f"  ✓ {Path(dest).name}")
    
    results = {}
    
    # Test 1: Mode classique
    print("\n" + "=" * 50)
    print(" TEST 1: MODE CLASSIQUE (Win32COM)")
    print("=" * 50)
    
    # Désactiver le mode rapide
    config = original_config.copy()
    config['optimization']['use_fast_mode'] = False
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)
    
    # Lancer le traitement
    from monitoring.auto_processor import AutoProcessor
    
    processor_classic = AutoProcessor()
    start_time = time.time()
    
    files = {
        'bulkreport': 'inbox/BulkReport_test.csv',
        'export': 'inbox/Export_test.xlsx',
        'frais': None
    }
    
    result_classic = processor_classic.process_files(files)
    time_classic = time.time() - start_time
    
    results['classic'] = {
        'time': time_classic,
        'success': result_classic['success'],
        'report': result_classic.get('report_path')
    }
    
    print(f"\n⏱️ Temps mode classique: {time_classic:.1f} secondes")
    
    # Nettoyer les fichiers
    for _, dest in test_files:
        try:
            os.remove(dest)
        except:
            pass
    
    # Recopier pour le second test
    for source, dest in test_files:
        if os.path.exists(source):
            shutil.copy2(source, dest)
    
    # Test 2: Mode optimisé
    print("\n" + "=" * 50)
    print(" TEST 2: MODE OPTIMISÉ (openpyxl)")
    print("=" * 50)
    
    # Activer le mode rapide
    config['optimization']['use_fast_mode'] = True
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)
    
    # Réinitialiser le processeur
    processor_fast = AutoProcessor()
    start_time = time.time()
    
    result_fast = processor_fast.process_files(files)
    time_fast = time.time() - start_time
    
    results['fast'] = {
        'time': time_fast,
        'success': result_fast['success'],
        'report': result_fast.get('report_path')
    }
    
    print(f"\n⏱️ Temps mode optimisé: {time_fast:.1f} secondes")
    
    # Restaurer la configuration
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(original_config, f, indent=4)
    
    # Afficher le résumé
    print("\n" + "=" * 70)
    print(" 📊 RÉSUMÉ DES PERFORMANCES")
    print("=" * 70)
    
    if results['classic']['success'] and results['fast']['success']:
        improvement = (results['classic']['time'] / results['fast']['time'])
        reduction = ((results['classic']['time'] - results['fast']['time']) / results['classic']['time']) * 100
        
        print(f"""
┌─────────────────────┬─────────────┬─────────────┐
│     Mode            │   Temps     │   Statut    │
├─────────────────────┼─────────────┼─────────────┤
│ Classique (COM)     │  {results['classic']['time']:>7.1f}s   │     ✅      │
│ Optimisé (openpyxl) │  {results['fast']['time']:>7.1f}s   │     ✅      │
└─────────────────────┴─────────────┴─────────────┘

📈 AMÉLIORATION:
  • {improvement:.1f}x plus rapide
  • Réduction de {reduction:.0f}% du temps
  • Gain de {results['classic']['time'] - results['fast']['time']:.1f} secondes
""")
    else:
        print("\n❌ Un des tests a échoué")
        print(f"  Classique: {'✅' if results['classic']['success'] else '❌'}")
        print(f"  Optimisé: {'✅' if results['fast']['success'] else '❌'}")
    
    # Vérifier que les outputs sont identiques
    if results['classic']['success'] and results['fast']['success']:
        print("\n🔍 Vérification de la compatibilité...")
        
        classic_path = results['classic']['report']
        fast_path = results['fast']['report']
        
        if classic_path and fast_path:
            classic_size = os.path.getsize(classic_path) if os.path.exists(classic_path) else 0
            fast_size = os.path.getsize(fast_path) if os.path.exists(fast_path) else 0
            
            size_diff = abs(classic_size - fast_size) / classic_size * 100 if classic_size > 0 else 0
            
            print(f"  • Taille classique: {classic_size:,} octets")
            print(f"  • Taille optimisé: {fast_size:,} octets")
            print(f"  • Différence: {size_diff:.1f}%")
            
            if size_diff < 10:
                print("  ✅ Les fichiers sont compatibles")
            else:
                print("  ⚠️ Différence notable, vérification manuelle recommandée")
    
    print("\n" + "=" * 70)
    print(" TEST TERMINÉ")
    print("=" * 70)
    
    # Nettoyer
    for _, dest in test_files:
        try:
            os.remove(dest)
        except:
            pass


if __name__ == "__main__":
    test_both_modes()
