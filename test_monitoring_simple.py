"""
Test simple du monitoring SANS configuration email
"""
import os
import shutil
from pathlib import Path

print("""
========================================
 TEST MONITORING (SANS EMAIL)
========================================

Ce test va:
1. Créer les dossiers nécessaires
2. Copier les fichiers de test
3. Lancer le monitoring
4. Générer automatiquement:
   - Rapport Excel
   - PDF
   - Archivage

SANS envoyer d'email (désactivé)
""")

# Créer les dossiers
folders = ['inbox', 'processed', 'errors', 'logs', 'outputs']
for folder in folders:
    Path(folder).mkdir(exist_ok=True)
    print(f"✓ {folder}/")

# Copier les fichiers de test si disponibles
test_files = [
    ('test_data/BulkReport_Test.csv', 'inbox/BulkReport_test.csv'),
    ('test_data/Export_Test.xlsx', 'inbox/Export_test.xlsx')
]

print("\nCopie des fichiers de test:")
for source, dest in test_files:
    if os.path.exists(source):
        shutil.copy2(source, dest)
        print(f"  ✓ {Path(dest).name}")

print("\n" + "="*40)
print("INSTRUCTIONS:")
print("="*40)
print("""
1. Ouvrez un nouveau terminal

2. Lancez le monitoring:
   python monitoring/auto_processor.py

3. Le système va automatiquement:
   - Détecter les fichiers
   - Générer le rapport Excel
   - Créer le PDF
   - Archiver les fichiers

4. Vérifiez les résultats dans:
   - outputs/ (rapports générés)
   - processed/ (fichiers archivés)

Note: L'envoi d'email est désactivé
""")

print(f"\nDossier à surveiller: {Path('inbox').absolute()}")
print("\nPrêt à démarrer!")
