"""
Système d'envoi d'emails professionnel avec templates HTML et pièces jointes
"""
import os
import smtplib
import logging
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.utils import formatdate, make_msgid
from typing import List, Optional, Dict
import json
from datetime import datetime
import base64

logger = logging.getLogger(__name__)


class ProfessionalEmailSender:
    """Gestionnaire d'envoi d'emails avec templates et tracking"""
    
    def __init__(self, config_path: str = "config/email_config.json"):
        """
        Initialise le système d'email
        
        Args:
            config_path: Chemin vers la configuration email
        """
        self.config = self._load_config(config_path)
        self.templates = self._load_templates()
        self.email_stats = {
            'sent': 0,
            'failed': 0,
            'last_sent': None
        }
    
    def _load_config(self, config_path: str) -> dict:
        """Charge la configuration email"""
        default_config = {
            'smtp': {
                'server': 'smtp.gmail.com',
                'port': 587,
                'use_tls': True,
                'username': 'your_email@gmail.com',
                'password': 'your_app_password'
            },
            'sender': {
                'name': 'UGP Reporter System',
                'email': 'ugp.reporter@gmail.com'
            },
            'partners': [
                {
                    'name': 'Responsable UGP',
                    'email': 'responsable@ugp.td',
                    'cc': [],
                    'send_pdf': True,
                    'send_excel': False
                }
            ],
            'templates': {
                'report_ready': 'email_templates/report_ready.html',
                'error_notification': 'email_templates/error.html',
                'weekly_summary': 'email_templates/weekly.html'
            },
            'branding': {
                'logo_path': 'assets/logo.png',
                'primary_color': '#2E7D32',
                'secondary_color': '#66BB6A'
            }
        }
        
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    # Fusionner les configs
                    self._merge_configs(default_config, user_config)
        except Exception as e:
            logger.warning(f"⚠️ Utilisation config email par défaut: {e}")
        
        return default_config
    
    def _merge_configs(self, default: dict, user: dict):
        """Fusionne les configurations de manière récursive"""
        for key, value in user.items():
            if key in default and isinstance(default[key], dict) and isinstance(value, dict):
                self._merge_configs(default[key], value)
            else:
                default[key] = value
    
    def _load_templates(self) -> dict:
        """Charge les templates HTML"""
        templates = {}
        
        # Template par défaut si fichiers non trouvés
        default_template = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
                .container { max-width: 600px; margin: 0 auto; padding: 20px; }
                .header { background: linear-gradient(135deg, #2E7D32, #66BB6A); color: white; padding: 30px; text-align: center; border-radius: 10px 10px 0 0; }
                .content { background: #f9f9f9; padding: 30px; border: 1px solid #ddd; border-top: none; }
                .footer { background: #333; color: white; padding: 20px; text-align: center; font-size: 12px; border-radius: 0 0 10px 10px; }
                .button { display: inline-block; padding: 12px 30px; background: #2E7D32; color: white; text-decoration: none; border-radius: 5px; margin: 20px 0; }
                .stats { background: white; padding: 20px; border-radius: 5px; margin: 20px 0; }
                .stats-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #eee; }
                .highlight { color: #2E7D32; font-weight: bold; }
                h1 { margin: 0; }
                h2 { color: #2E7D32; }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>{{title}}</h1>
                    <p>{{subtitle}}</p>
                </div>
                <div class="content">
                    {{content}}
                </div>
                <div class="footer">
                    <p>© 2025 UGP Reporter - Système Automatique de Rapports</p>
                    <p>Ce message a été généré automatiquement. Ne pas répondre.</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        templates['default'] = default_template
        
        # Template rapport prêt
        templates['report_ready'] = default_template.replace(
            '{{content}}',
            """
            <h2>Rapport de Paiement Généré avec Succès ✅</h2>
            <p>Bonjour {{recipient_name}},</p>
            <p>Le rapport de paiement a été généré automatiquement et est disponible en pièce jointe.</p>
            
            <div class="stats">
                <h3>📊 Détails du Rapport</h3>
                <div class="stats-row">
                    <span>Date de génération:</span>
                    <span class="highlight">{{generation_date}}</span>
                </div>
                <div class="stats-row">
                    <span>Nombre de transactions:</span>
                    <span class="highlight">{{transaction_count}}</span>
                </div>
                <div class="stats-row">
                    <span>Montant total:</span>
                    <span class="highlight">{{total_amount}} FCFA</span>
                </div>
                <div class="stats-row">
                    <span>Frais totaux:</span>
                    <span class="highlight">{{total_fees}} FCFA</span>
                </div>
                <div class="stats-row">
                    <span>Bénéficiaires uniques:</span>
                    <span class="highlight">{{unique_beneficiaries}}</span>
                </div>
            </div>
            
            <p><strong>📎 Pièces jointes:</strong></p>
            <ul>
                <li>Rapport PDF pour impression</li>
                <li>Rapport Excel (si demandé)</li>
            </ul>
            
            <p>Pour toute question, veuillez contacter le support technique.</p>
            
            <center>
                <a href="#" class="button">Voir le Dashboard</a>
            </center>
            """
        )
        
        # Template d'erreur
        templates['error'] = default_template.replace(
            '{{content}}',
            """
            <h2 style="color: #d32f2f;">⚠️ Erreur lors du Traitement</h2>
            <p>Bonjour {{recipient_name}},</p>
            <p>Une erreur s'est produite lors du traitement automatique des fichiers.</p>
            
            <div style="background: #ffebee; padding: 15px; border-radius: 5px; margin: 20px 0;">
                <strong>Détails de l'erreur:</strong><br>
                {{error_message}}
            </div>
            
            <p><strong>Fichiers concernés:</strong></p>
            <ul>
                {{file_list}}
            </ul>
            
            <p>Les fichiers ont été déplacés dans le dossier d'erreur pour vérification manuelle.</p>
            """
        )
        
        return templates
    
    def send_report_email(self, recipient: Dict, report_data: Dict, 
                          attachments: List[str] = None) -> bool:
        """
        Envoie un email avec le rapport en pièce jointe
        
        Args:
            recipient: Dictionnaire avec infos du destinataire
            report_data: Données du rapport pour le template
            attachments: Liste des chemins de fichiers à joindre
        
        Returns:
            True si succès, False sinon
        """
        try:
            # Créer le message
            msg = MIMEMultipart('related')
            msg['From'] = f"{self.config['sender']['name']} <{self.config['sender']['email']}>"
            msg['To'] = recipient['email']
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = f"📊 Rapport UGP - {report_data.get('date', datetime.now().strftime('%d/%m/%Y'))}"
            msg['Message-ID'] = make_msgid()
            
            # Ajouter CC si spécifié
            if recipient.get('cc'):
                msg['Cc'] = ', '.join(recipient['cc'])
            
            # Préparer le contenu HTML
            html_content = self._render_template('report_ready', {
                'title': 'Rapport de Paiement UGP',
                'subtitle': f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
                'recipient_name': recipient.get('name', 'Partenaire'),
                'generation_date': datetime.now().strftime('%d/%m/%Y %H:%M'),
                'transaction_count': report_data.get('transaction_count', 0),
                'total_amount': f"{report_data.get('total_amount', 0):,.0f}".replace(',', ' '),
                'total_fees': f"{report_data.get('total_fees', 0):,.0f}".replace(',', ' '),
                'unique_beneficiaries': report_data.get('unique_beneficiaries', 0)
            })
            
            # Attacher le HTML
            msg_html = MIMEText(html_content, 'html', 'utf-8')
            msg.attach(msg_html)
            
            # Ajouter les pièces jointes
            if attachments:
                for file_path in attachments:
                    if os.path.exists(file_path):
                        self._attach_file(msg, file_path)
            
            # Envoyer l'email
            success = self._send_email(msg, recipient['email'])
            
            if success:
                self.email_stats['sent'] += 1
                self.email_stats['last_sent'] = datetime.now()
                logger.info(f"✅ Email envoyé à: {recipient['email']}")
            else:
                self.email_stats['failed'] += 1
                
            return success
            
        except Exception as e:
            logger.error(f"❌ Erreur envoi email: {e}")
            self.email_stats['failed'] += 1
            return False
    
    def _render_template(self, template_name: str, data: Dict) -> str:
        """Rend un template avec les données"""
        template = self.templates.get(template_name, self.templates['default'])
        
        # Remplacer les variables
        for key, value in data.items():
            template = template.replace(f'{{{{{key}}}}}', str(value))
        
        return template
    
    def _attach_file(self, msg: MIMEMultipart, file_path: str):
        """Attache un fichier au message"""
        try:
            file_path = Path(file_path)
            
            with open(file_path, 'rb') as f:
                attach = MIMEApplication(f.read())
                attach.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=file_path.name
                )
                msg.attach(attach)
            
            logger.info(f"  📎 Pièce jointe: {file_path.name}")
            
        except Exception as e:
            logger.warning(f"  ⚠️ Impossible d'attacher: {file_path.name} - {e}")
    
    def _send_email(self, msg: MIMEMultipart, recipient: str) -> bool:
        """Envoie effectivement l'email via SMTP"""
        try:
            # Connexion SMTP
            smtp_config = self.config['smtp']
            
            with smtplib.SMTP(smtp_config['server'], smtp_config['port']) as server:
                if smtp_config['use_tls']:
                    server.starttls()
                
                # Authentification
                server.login(smtp_config['username'], smtp_config['password'])
                
                # Envoi
                server.send_message(msg)
                
                return True
                
        except Exception as e:
            logger.error(f"❌ Erreur SMTP: {e}")
            return False
    
    def send_to_all_partners(self, report_data: Dict, attachments: List[str]) -> Dict:
        """
        Envoie le rapport à tous les partenaires configurés
        
        Args:
            report_data: Données du rapport
            attachments: Liste des pièces jointes
        
        Returns:
            Dictionnaire avec les résultats d'envoi
        """
        results = {
            'success': [],
            'failed': []
        }
        
        partners = self.config.get('partners', [])
        logger.info(f"📧 Envoi à {len(partners)} partenaires...")
        
        for partner in partners:
            # Filtrer les pièces jointes selon les préférences
            partner_attachments = []
            for attachment in attachments:
                if '.pdf' in attachment and partner.get('send_pdf', True):
                    partner_attachments.append(attachment)
                elif '.xlsx' in attachment and partner.get('send_excel', False):
                    partner_attachments.append(attachment)
            
            # Envoyer
            success = self.send_report_email(partner, report_data, partner_attachments)
            
            if success:
                results['success'].append(partner['email'])
            else:
                results['failed'].append(partner['email'])
        
        # Résumé
        logger.info(f"📊 Résultat: {len(results['success'])} succès, {len(results['failed'])} échecs")
        
        return results
    
    def get_stats(self) -> Dict:
        """Retourne les statistiques d'envoi"""
        return {
            'total_sent': self.email_stats['sent'],
            'total_failed': self.email_stats['failed'],
            'last_sent': self.email_stats['last_sent'].isoformat() if self.email_stats['last_sent'] else None,
            'success_rate': (
                self.email_stats['sent'] / (self.email_stats['sent'] + self.email_stats['failed']) * 100
                if (self.email_stats['sent'] + self.email_stats['failed']) > 0 else 0
            )
        }
