"""
Système de monitoring intelligent pour surveillance automatique des dossiers
"""
import os
import time
import logging
from pathlib import Path
from datetime import datetime, timedelta
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import json
import threading
from typing import Dict, List, Optional
import hashlib

logger = logging.getLogger(__name__)


class SmartFileWatcher(FileSystemEventHandler):
    """Surveillant intelligent de dossiers avec détection de patterns"""
    
    def __init__(self, config_path: str = "config/monitoring_config.json"):
        """
        Initialise le système de monitoring
        
        Args:
            config_path: Chemin vers la configuration du monitoring
        """
        self.config = self._load_config(config_path)
        self.watched_folder = Path(self.config['watched_folder'])
        self.processed_folder = Path(self.config['processed_folder'])
        self.error_folder = Path(self.config['error_folder'])
        
        # Créer les dossiers s'ils n'existent pas
        self._create_folders()
        
        # État du monitoring
        self.pending_files = {}
        self.processing_queue = []
        self.file_checksums = {}
        self.last_check = datetime.now()
        
        # Patterns de fichiers requis
        self.required_patterns = {
            'bulkreport': self.config['patterns']['bulkreport'],
            'export': self.config['patterns']['export'],
            'frais': self.config['patterns'].get('frais', ['frais'])
        }
        
        # Callback pour traitement
        self.process_callback = None
        
        # Thread de vérification périodique
        self.check_thread = threading.Thread(target=self._periodic_check, daemon=True)
        self.check_thread.start()
        
        logger.info(f"🔍 Monitoring initialisé sur: {self.watched_folder}")
    
    def _load_config(self, config_path: str) -> dict:
        """Charge la configuration du monitoring"""
        default_config = {
            'watched_folder': './inbox',
            'processed_folder': './processed',
            'error_folder': './errors',
            'check_interval': 5,
            'file_stability_time': 2,
            'patterns': {
                'bulkreport': ['bulkreport', 'bulk'],
                'export': ['export', 'beneficiaire'],
                'frais': ['frais', 'fee', 'commission']
            },
            'auto_process': True,
            'archive_processed': True,
            'send_notifications': True
        }
        
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    default_config.update(user_config)
        except Exception as e:
            logger.warning(f"⚠️ Configuration par défaut utilisée: {e}")
        
        return default_config
    
    def _create_folders(self):
        """Crée les dossiers nécessaires"""
        for folder in [self.watched_folder, self.processed_folder, self.error_folder]:
            folder.mkdir(parents=True, exist_ok=True)
            logger.info(f"✓ Dossier vérifié: {folder}")
    
    def on_created(self, event):
        """Déclenché lors de la création d'un fichier"""
        if not event.is_directory:
            self._handle_new_file(event.src_path)
    
    def on_modified(self, event):
        """Déclenché lors de la modification d'un fichier"""
        if not event.is_directory:
            self._handle_new_file(event.src_path)
    
    def _handle_new_file(self, file_path: str):
        """Gère l'arrivée d'un nouveau fichier"""
        file_path = Path(file_path)
        
        # Ignorer les fichiers temporaires
        if file_path.name.startswith('~') or file_path.name.startswith('.'):
            return
        
        # Vérifier si le fichier est stable (pas en cours d'écriture)
        if not self._is_file_stable(file_path):
            return
        
        logger.info(f"📄 Nouveau fichier détecté: {file_path.name}")
        
        # Identifier le type de fichier
        file_type = self._identify_file_type(file_path)
        if file_type:
            self.pending_files[file_type] = {
                'path': file_path,
                'timestamp': datetime.now(),
                'checksum': self._calculate_checksum(file_path)
            }
            logger.info(f"  → Identifié comme: {file_type}")
            
            # Vérifier si on a tous les fichiers requis
            self._check_complete_set()
    
    def _is_file_stable(self, file_path: Path) -> bool:
        """Vérifie si un fichier est stable (fini d'être écrit)"""
        try:
            # Attendre un peu
            time.sleep(self.config['file_stability_time'])
            
            # Vérifier que la taille n'a pas changé
            initial_size = file_path.stat().st_size
            time.sleep(0.5)
            final_size = file_path.stat().st_size
            
            return initial_size == final_size and initial_size > 0
        except:
            return False
    
    def _identify_file_type(self, file_path: Path) -> Optional[str]:
        """Identifie le type de fichier basé sur les patterns"""
        filename_lower = file_path.name.lower()
        
        for file_type, patterns in self.required_patterns.items():
            for pattern in patterns:
                if pattern.lower() in filename_lower:
                    return file_type
        
        return None
    
    def _calculate_checksum(self, file_path: Path) -> str:
        """Calcule le checksum SHA256 d'un fichier"""
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    
    def _check_complete_set(self):
        """Vérifie si on a un ensemble complet de fichiers"""
        required = ['bulkreport', 'export']  # 'frais' est optionnel
        
        if all(req in self.pending_files for req in required):
            logger.info("✅ Ensemble complet détecté! Lancement du traitement...")
            
            # Préparer les fichiers pour traitement
            files_to_process = {
                'bulkreport': str(self.pending_files['bulkreport']['path']),
                'export': str(self.pending_files['export']['path']),
                'frais': str(self.pending_files['frais']['path']) if 'frais' in self.pending_files else None
            }
            
            # Ajouter à la queue de traitement
            self.processing_queue.append({
                'id': datetime.now().strftime('%Y%m%d_%H%M%S'),
                'files': files_to_process,
                'timestamp': datetime.now()
            })
            
            # Déclencher le callback de traitement
            if self.process_callback:
                threading.Thread(
                    target=self._process_with_callback,
                    args=(files_to_process,),
                    daemon=True
                ).start()
            
            # Nettoyer les fichiers pending
            self.pending_files.clear()
    
    def _process_with_callback(self, files: Dict[str, str]):
        """Exécute le callback de traitement avec gestion d'erreur"""
        try:
            logger.info("🔄 Début du traitement automatique...")
            result = self.process_callback(files)
            
            if result['success']:
                logger.info("✅ Traitement réussi!")
                self._archive_processed_files(files)
            else:
                logger.error(f"❌ Erreur de traitement: {result.get('error')}")
                self._move_to_error_folder(files)
                
        except Exception as e:
            logger.error(f"❌ Exception lors du traitement: {e}")
            self._move_to_error_folder(files)
    
    def _archive_processed_files(self, files: Dict[str, str]):
        """Archive les fichiers traités avec succès"""
        if not self.config['archive_processed']:
            return
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        archive_folder = self.processed_folder / timestamp
        archive_folder.mkdir(exist_ok=True)
        
        for file_type, file_path in files.items():
            if file_path and os.path.exists(file_path):
                source = Path(file_path)
                dest = archive_folder / source.name
                source.rename(dest)
                logger.info(f"  📦 Archivé: {source.name} → {archive_folder.name}/")
    
    def _move_to_error_folder(self, files: Dict[str, str]):
        """Déplace les fichiers en erreur"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        error_folder = self.error_folder / timestamp
        error_folder.mkdir(exist_ok=True)
        
        for file_type, file_path in files.items():
            if file_path and os.path.exists(file_path):
                source = Path(file_path)
                dest = error_folder / source.name
                source.rename(dest)
                logger.warning(f"  ⚠️ Déplacé en erreur: {source.name}")
    
    def _periodic_check(self):
        """Vérification périodique des fichiers orphelins"""
        while True:
            time.sleep(self.config['check_interval'])
            
            # Nettoyer les fichiers pending trop vieux (> 1 heure)
            cutoff_time = datetime.now() - timedelta(hours=1)
            
            for file_type in list(self.pending_files.keys()):
                if self.pending_files[file_type]['timestamp'] < cutoff_time:
                    logger.info(f"🧹 Nettoyage fichier orphelin: {file_type}")
                    del self.pending_files[file_type]
    
    def set_process_callback(self, callback):
        """Définit la fonction callback pour le traitement"""
        self.process_callback = callback
        logger.info("✓ Callback de traitement configuré")
    
    def start_monitoring(self):
        """Démarre le monitoring du dossier"""
        observer = Observer()
        observer.schedule(self, str(self.watched_folder), recursive=False)
        observer.start()
        
        logger.info(f"👁️ Monitoring actif sur: {self.watched_folder}")
        logger.info("  → En attente de fichiers...")
        
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            observer.stop()
            logger.info("⏹️ Monitoring arrêté")
        
        observer.join()
    
    def get_stats(self) -> dict:
        """Retourne les statistiques du monitoring"""
        return {
            'watched_folder': str(self.watched_folder),
            'pending_files': len(self.pending_files),
            'queue_size': len(self.processing_queue),
            'last_check': self.last_check.isoformat(),
            'status': 'running'
        }
