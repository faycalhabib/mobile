"""
Script de test pour le système de monitoring automatique
"""
import os
import shutil
from pathlib import Path
import time

# Créer les dossiers nécessaires
folders = ['inbox', 'processed', 'errors', 'logs', 'outputs']
for folder in folders:
    Path(folder).mkdir(exist_ok=True)
    print(f"✓ Dossier créé/vérifié: {folder}")

print("\n" + "="*60)
print(" TEST DU SYSTÈME DE MONITORING")
print("="*60)

print("""
Instructions pour tester:

1. Ouvrez un nouveau terminal
2. Lancez: python monitoring/auto_processor.py

3. Dans le dossier 'inbox', copiez:
   - Un fichier BulkReport.csv
   - Un fichier Export.xlsx
   - (Optionnel) Un fichier Frais.xlsx

4. Le système devrait:
   ✓ Détecter automatiquement les fichiers
   ✓ Générer le rapport Excel
   ✓ Convertir en PDF
   ✓ Envoyer par email (si configuré)
   ✓ Archiver les fichiers dans 'processed'

5. En cas d'erreur, les fichiers seront dans 'errors'

Dossier inbox: {}
""".format(Path('inbox').absolute()))

# Copier les fichiers de test si disponibles
test_files = {
    'bulk': 'test_data/BulkReport_Test.csv',
    'export': 'test_data/Export_Test.xlsx'
}

print("\nCopie des fichiers de test...")
for file_type, source in test_files.items():
    if os.path.exists(source):
        dest_name = f"Test_{file_type}_{Path(source).name}"
        dest = Path('inbox') / dest_name
        shutil.copy2(source, dest)
        print(f"  ✓ Copié: {dest_name}")
        
        # Renommer pour que le monitoring les détecte
        if 'bulk' in file_type:
            new_name = dest.parent / "BulkReport_test.csv"
            dest.rename(new_name)
            print(f"    → Renommé en: BulkReport_test.csv")
        elif 'export' in file_type:
            new_name = dest.parent / "Export_test.xlsx"
            dest.rename(new_name)
            print(f"    → Renommé en: Export_test.xlsx")

print("\n✅ Fichiers de test prêts dans 'inbox'")
print("🚀 Lancez maintenant: python monitoring/auto_processor.py")
