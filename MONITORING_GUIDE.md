# 🚀 GUIDE DU SYSTÈME DE MONITORING AUTOMATIQUE

## 📋 Vue d'ensemble

Le système de monitoring automatique surveille un dossier et traite automatiquement les fichiers déposés:

```
📁 inbox/ (surveillé)
    ↓
🔄 Traitement automatique
    ↓
📊 Rapport Excel
    ↓
📄 Conversion PDF
    ↓
📧 Envoi Email
    ↓
📦 Archivage dans processed/
```

## 🎯 Fonctionnalités

### ✅ **Monitoring Intelligent**
- Surveillance en temps réel du dossier `inbox/`
- Détection automatique des patterns de fichiers
- Vérification de stabilité des fichiers (anti-corruption)
- File d'attente avec priorités
- Retry automatique en cas d'échec

### 📄 **Conversion PDF Professionnelle**
- Mise en page optimisée pour impression
- Qualité configurable (minimum/standard/maximum)
- En-têtes et pieds de page personnalisés
- Centrage automatique
- Ajustement à la page

### 📧 **Envoi Email Automatique**
- Templates HTML professionnels
- Support multi-destinataires avec CC
- Pièces jointes (PDF + Excel optionnel)
- Statistiques dans le corps du mail
- Notification d'erreur aux admins

### 🔐 **Sécurité & Robustesse**
- Checksum SHA256 pour intégrité
- Archivage automatique des fichiers traités
- Isolation des erreurs dans dossier séparé
- Logs détaillés pour audit
- Configuration externalisée (JSON)

## 📦 Installation

### 1. Installer les dépendances
```bash
pip install watchdog python-dotenv
```

### 2. Créer les dossiers
```bash
mkdir inbox processed errors logs outputs
```

### 3. Configurer les emails
Éditer `config/email_config.json`:
```json
{
    "smtp": {
        "server": "smtp.gmail.com",
        "port": 587,
        "username": "votre.email@gmail.com",
        "password": "mot_de_passe_application"
    },
    "partners": [
        {
            "name": "Responsable",
            "email": "responsable@entreprise.com",
            "send_pdf": true
        }
    ]
}
```

**Note Gmail**: Utilisez un [mot de passe d'application](https://support.google.com/accounts/answer/185833)

## 🚀 Utilisation

### Méthode 1: Script Batch (Windows)
```bash
start_monitoring.bat
```

### Méthode 2: Python Direct
```bash
python monitoring/auto_processor.py
```

### Méthode 3: Test avec fichiers exemples
```bash
python test_monitoring.py
python monitoring/auto_processor.py
```

## 📁 Structure des dossiers

```
UGP_Reporter/
├── inbox/              # 👈 Déposez vos fichiers ICI
│   ├── BulkReport.csv
│   ├── Export.xlsx
│   └── Frais.xlsx (optionnel)
├── processed/          # ✅ Fichiers traités avec succès
│   └── 20250117_143022/
│       ├── BulkReport.csv
│       ├── Export.xlsx
│       └── Frais.xlsx
├── errors/             # ❌ Fichiers en erreur
│   └── 20250117_143023/
│       └── fichiers...
├── outputs/            # 📊 Rapports générés
│   ├── Rapport_AUTO_20250117_143022.xlsx
│   └── Rapport_AUTO_20250117_143022.pdf
└── logs/               # 📝 Fichiers de log
    └── auto_processor.log
```

## ⚙️ Configuration Avancée

### Patterns de détection (`config/monitoring_config.json`)
```json
{
    "patterns": {
        "bulkreport": ["bulkreport", "bulk", "rapport"],
        "export": ["export", "beneficiaire", "etat"],
        "frais": ["frais", "fee", "tarif"]
    }
}
```

### Options PDF (`config/auto_processor_config.json`)
```json
{
    "pdf_options": {
        "quality": "standard",
        "orientation": "portrait",
        "fit_to_page": true,
        "margins": "normal"
    }
}
```

### Multi-destinataires Email
```json
{
    "partners": [
        {
            "name": "Directeur",
            "email": "directeur@ugp.td",
            "cc": ["assistant@ugp.td"],
            "send_pdf": true,
            "send_excel": true
        },
        {
            "name": "Comptabilité",
            "email": "compta@ugp.td",
            "send_pdf": true,
            "send_excel": false
        }
    ]
}
```

## 📊 Workflow Complet

1. **Détection**: Le système détecte les nouveaux fichiers
2. **Validation**: Vérification de la stabilité (fichier complet)
3. **Identification**: Reconnaissance du type (BulkReport/Export/Frais)
4. **Attente**: Attente de l'ensemble complet (minimum 2 fichiers)
5. **Traitement**: Génération du rapport Excel
6. **Conversion**: Création du PDF professionnel
7. **Notification**: Envoi email avec pièces jointes
8. **Archivage**: Déplacement dans `processed/`
9. **Logging**: Enregistrement de toutes les opérations

## 🔍 Monitoring en temps réel

Le système affiche en temps réel:
```
👁️ Monitoring actif sur: ./inbox
  → En attente de fichiers...

📄 Nouveau fichier détecté: BulkReport_130809.csv
  → Identifié comme: bulkreport
📄 Nouveau fichier détecté: Export_0131.xlsx
  → Identifié comme: export
✅ Ensemble complet détecté! Lancement du traitement...

🔄 Début du traitement automatique...
  ✓ BulkReport: 127 transactions
  ✓ Export: 127 bénéficiaires
  ✓ 127 transactions traitées
  ✓ Rapport généré: Rapport_AUTO_20250117_143022.xlsx
  ✓ PDF généré: Rapport_AUTO_20250117_143022.pdf
  ✓ Emails envoyés: 3 destinataires
✅ TRAITEMENT TERMINÉ AVEC SUCCÈS

📦 Archivé: BulkReport_130809.csv → processed/20250117_143022/
📦 Archivé: Export_0131.xlsx → processed/20250117_143022/
```

## 🛠️ Dépannage

### Problème: "Fichiers non détectés"
- Vérifiez les patterns dans `monitoring_config.json`
- Le nom doit contenir un des mots-clés (bulkreport, export, etc.)

### Problème: "Erreur de conversion PDF"
- Vérifiez que Excel est installé
- Fermez tous les fichiers Excel ouverts

### Problème: "Email non envoyé"
- Vérifiez `config/email_config.json`
- Pour Gmail: activez l'accès moins sécurisé ou utilisez un mot de passe d'application
- Vérifiez la connexion internet

### Problème: "Permission denied"
- Fermez les fichiers ouverts dans Excel
- Vérifiez les permissions des dossiers

## 🚨 Logs et Débogage

Consultez les logs détaillés:
```bash
tail -f logs/auto_processor.log
```

Niveau de log configurable:
- INFO: Opérations normales
- WARNING: Avertissements
- ERROR: Erreurs
- DEBUG: Débogage détaillé

## 📈 Statistiques

Le système maintient des statistiques:
- Nombre total de traitements
- Taux de succès
- Dernière exécution
- Temps moyen de traitement
- Emails envoyés/échoués
- Conversions PDF réussies

## 🔄 Intégrations Futures

- **API REST**: Endpoint pour déclencher manuellement
- **Dashboard Web**: Interface de monitoring temps réel
- **Webhook**: Notifications Slack/Teams
- **Cloud Storage**: Backup automatique sur Google Drive
- **Machine Learning**: Détection d'anomalies
- **Scheduler**: Traitement programmé (CRON)

## 💡 Tips Pro

1. **Performance**: Gardez le dossier `inbox/` propre
2. **Sécurité**: Utilisez des mots de passe d'application
3. **Backup**: Activez l'archivage automatique
4. **Monitoring**: Consultez régulièrement les logs
5. **Test**: Utilisez d'abord les fichiers de test

## 📞 Support

En cas de problème:
1. Consultez les logs dans `logs/auto_processor.log`
2. Vérifiez la configuration dans `config/`
3. Testez avec les fichiers d'exemple
4. Contactez le support technique

---

*Développé avec ❤️ pour UGP - Union des Groupements de Producteurs*
