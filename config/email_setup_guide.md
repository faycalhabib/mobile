# 📧 GUIDE DE CONFIGURATION EMAIL

## ⚠️ IMPORTANT : Configuration requise pour l'envoi d'emails

### 1. Configuration Gmail (Recommandé)

#### Étape 1 : Activer la vérification en 2 étapes
1. Allez sur https://myaccount.google.com/security
2. Activez la "Vérification en 2 étapes"

#### Étape 2 : Créer un mot de passe d'application
1. Allez sur https://myaccount.google.com/apppasswords
2. Sélectionnez "Mail" et "Ordinateur Windows"
3. Cliquez sur "Générer"
4. **IMPORTANT** : Copiez le mot de passe généré (16 caractères)

#### Étape 3 : Configurer email_config.json
```json
{
    "smtp": {
        "server": "smtp.gmail.com",
        "port": 587,
        "use_tls": true,
        "username": "votre.email@gmail.com",
        "password": "xxxx xxxx xxxx xxxx"  // Mot de passe d'application (sans espaces)
    }
}
```

### 2. Configuration Outlook/Hotmail

```json
{
    "smtp": {
        "server": "smtp.office365.com",
        "port": 587,
        "use_tls": true,
        "username": "votre.email@outlook.com",
        "password": "votre_mot_de_passe"
    }
}
```

### 3. Configuration Yahoo

```json
{
    "smtp": {
        "server": "smtp.mail.yahoo.com",
        "port": 587,
        "use_tls": true,
        "username": "votre.email@yahoo.com",
        "password": "mot_de_passe_application"
    }
}
```

## 🔒 Sécurité

### Option 1 : Variables d'environnement (Recommandé)
Créez un fichier `.env` :
```
EMAIL_USERNAME=votre.email@gmail.com
EMAIL_PASSWORD=xxxx xxxx xxxx xxxx
```

### Option 2 : Désactiver temporairement l'envoi d'emails
Dans `config/auto_processor_config.json` :
```json
{
    "processing": {
        "send_email": false  // Mettre à false pour désactiver
    }
}
```

## 🎯 Configuration des destinataires

Dans `config/email_config.json`, section `partners` :

```json
{
    "partners": [
        {
            "name": "Nom du Responsable",
            "email": "email@entreprise.com",
            "cc": ["copie1@entreprise.com", "copie2@entreprise.com"],
            "send_pdf": true,   // Envoyer le PDF
            "send_excel": false  // Ne pas envoyer l'Excel
        }
    ]
}
```

## ✅ Test de configuration

Pour tester votre configuration email :

1. Configurez d'abord `email_config.json` avec vos identifiants
2. Lancez le test :
```python
python test_email_config.py
```

## ❌ Résolution des problèmes

### Erreur : "Username and Password not accepted"
- Vérifiez que vous utilisez un mot de passe d'application (pas votre mot de passe normal)
- Pour Gmail : https://support.google.com/accounts/answer/185833

### Erreur : "Connection refused"
- Vérifiez votre connexion internet
- Vérifiez que le port n'est pas bloqué par un firewall
- Essayez le port 465 avec SSL au lieu de 587 avec TLS

### Erreur : "Timeout"
- Augmentez le timeout dans le code
- Vérifiez les paramètres du serveur SMTP

## 💡 Conseils

1. **Ne jamais** commiter vos mots de passe sur GitHub
2. Utilisez toujours des mots de passe d'application
3. Testez d'abord avec un seul destinataire
4. Vérifiez les limites d'envoi de votre fournisseur email

## 🚀 Démarrage rapide sans email

Si vous voulez tester le système sans configurer les emails :

1. Désactivez l'envoi dans `config/auto_processor_config.json`
2. Le système fonctionnera normalement (monitoring, génération de rapports, PDF)
3. Seul l'envoi d'email sera ignoré
