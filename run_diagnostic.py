"""
Lancer le diagnostic complet du système
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.full_diagnostic import FullDiagnostic
import shutil

# Chemins des fichiers
bulk_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\BulkReport_130809.csv"
export_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Export_0131-FMC19-Beat.xlsx"
fees_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\frais.xlsx"
template_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx"
output_path = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP_Reporter\outputs\diagnostic_test.xlsx"

# Créer le dossier de sortie
os.makedirs(os.path.dirname(output_path), exist_ok=True)

# Copier le template
if os.path.exists(template_path):
    shutil.copy2(template_path, output_path)
    print(f"✓ Template copié vers: {output_path}")
else:
    print(f"✗ Template non trouvé: {template_path}")
    sys.exit(1)

# Lancer le diagnostic
diagnostic = FullDiagnostic()
results = diagnostic.scan_full_process(bulk_path, export_path, fees_path, output_path)

# Afficher le chemin du fichier de sortie
print(f"\n📄 Fichier de test: {output_path}")
print("Ouvrez ce fichier pour vérifier si les données sont présentes")
