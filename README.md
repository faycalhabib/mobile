# 📊 UGP Reporter - Générateur de Rapports Automatique

Application desktop moderne# UGP Reporter 📊

Système automatisé de génération de rapports de paiement pour UGP - Union des Groupements de Producteurs.

## 🎯 Fonctionnalités
- ✅ **Génération de rapports** Excel formatés
- ✅ **Interface moderne** avec thème sombre
- ✅ **Gestion d'erreurs robuste**
- ✅ **Cache des correspondances** pour optimisation

## 📋 Prérequis

- Windows 10/11 (ou Linux/Mac avec adaptations)
- Python 3.8 ou supérieur

## 🔧 Installation

### Option 1: Installation rapide (Windows)

1. Installer Python depuis [python.org](https://www.python.org/downloads/)
2. Double-cliquer sur `install.bat`
3. Lancer l'application avec `run.bat`

### Option 2: Installation manuelle

```bash
# Cloner ou télécharger le projet
cd UGP_Reporter

# Installer les dépendances
pip install -r requirements.txt

# Lancer l'application
python main.py
```

## 📁 Structure des fichiers

```
UGP_Reporter/
├── main.py                 # Application principale
├── core/
│   ├── file_handler.py     # Gestion des fichiers
│   ├── data_processor.py   # Traitement des données
│   └── report_generator.py # Génération des rapports
├── config/
│   └── settings.json       # Configuration
├── outputs/                # Rapports générés
└── logs/                   # Fichiers de log
```

## 🎯 Utilisation

### Étape 1: Préparer vos fichiers

Vous avez besoin de 3 fichiers:

1. **BulkReport.csv** - Données de paiement (téléchargé depuis le site)
2. **Export Excel** - Liste des bénéficiaires avec leurs noms
3. **Frais.xlsx** - Table de calcul des frais

### Étape 2: Lancer l'application

1. Ouvrir `UGP_Reporter.exe` ou lancer avec `python main.py`
2. Sélectionner les 3 fichiers requis
3. Vérifier/modifier les paramètres du rapport
4. Cliquer sur **"GÉNÉRER RAPPORT"**

### Étape 3: Récupérer le rapport

Le rapport sera généré dans le dossier `outputs/` avec le nom:
`Rapport_UGP_[DATE]_[HEURE].xlsx`

## 🎨 Format du rapport généré

Le rapport Excel contient:
- **En-tête** avec date, libellé, budget et projet
- **Tableau des transactions** avec:
  - Date et N° de transaction
  - Montant et frais calculés
  - Informations du bénéficiaire
- **Total** automatique en bas du tableau

## ⚙️ Configuration

Modifier `config/settings.json` pour personnaliser:
- Dossier de sortie
- Taux de frais par défaut
- Paramètres de l'interface

## 🐛 Résolution des problèmes

### Erreur "Module not found"
```bash
pip install --upgrade -r requirements.txt
```

### Fichier CSV non reconnu
- Vérifier l'encodage (UTF-8 recommandé)
- S'assurer que le fichier n'est pas corrompu

### Bénéficiaires non trouvés
- Vérifier que le fichier Export contient la colonne "Nom et prénoms"
- Le système utilisera "Bénéficiaire [numéro]" par défaut

## 📊 Exemple de workflow

1. **Input**: 
   - BulkReport_130809.csv (transactions)
   - Export_0131.xlsx (bénéficiaires)
   - frais.xlsx (table des frais)

2. **Processing**:
   - Extraction des transactions
   - Mapping téléphone → nom
   - Calcul des frais (1.68% par défaut)

3. **Output**:
   - Rapport_UGP_20250915_234800.xlsx

## 🔐 Sécurité

- Toutes les données restent en local
- Aucune connexion internet requise
- Logs stockés localement

## 📝 Licence

Propriété de l'entreprise UGP

## 💡 Support

Pour toute question ou problème, consulter les logs dans le dossier `logs/`

---

**Version**: 1.0.0  
**Dernière mise à jour**: Septembre 2025
