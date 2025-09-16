"""
Application principale - Générateur de Rapports UGP
Interface moderne avec CustomTkinter
"""
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import sys
import json
from datetime import datetime
import logging
import traceback
import subprocess
import platform

# Configuration des logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Importer les modules core
from core.file_handler import FileHandler
from core.data_processor import DataProcessor
from core.report_generator import ReportGenerator

# Configuration du thème
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class UGPReporterApp(ctk.CTk):
    """Application principale avec interface moderne"""
    
    def __init__(self):
        super().__init__()
        
        self.title("📊 Générateur de Rapports UGP")
        self.geometry("950x750")
        self.resizable(True, True)
        
        # Variables
        self.file_paths = {
            'bulk': tk.StringVar(),
            'export': tk.StringVar(),
            'fees': tk.StringVar(),
            'template': tk.StringVar(value=r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP\Rapport UGP.xlsx")
        }
        
        self.metadata = {
            'date_paiement': tk.StringVar(value=datetime.now().strftime("%d/%m/%Y")),
            'libelle': tk.StringVar(value="PAIEMENT LOCATION SALLE"),
            'budget': tk.StringVar(value="500000"),
            'projet': tk.StringVar(value="UGP")
        }
        
        self.processing = False
        self.errors = []
        
        # Charger la configuration
        self.load_config()
        
        # Créer les dossiers nécessaires
        os.makedirs("outputs", exist_ok=True)
        os.makedirs("logs", exist_ok=True)
        os.makedirs("config", exist_ok=True)
        
        # Construire l'interface
        self.setup_ui()
        
        # Centrer la fenêtre
        self.center_window()
    
    def center_window(self):
        """Centrer la fenêtre sur l'écran"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
    def load_config(self):
        """Charger la configuration sauvegardée"""
        try:
            with open("config/settings.json", 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Appliquer les valeurs par défaut
                defaults = config.get('defaults', {})
                self.metadata['projet'].set(defaults.get('projet', 'UGP'))
                self.metadata['budget'].set(str(defaults.get('budget', 500000)))
        except:
            pass
    
    def setup_ui(self):
        """Construire l'interface utilisateur"""
        # Container principal avec padding
        main_container = ctk.CTkFrame(self, corner_radius=0)
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Titre
        title_frame = ctk.CTkFrame(main_container, height=60)
        title_frame.pack(fill="x", pady=(0, 20))
        title_frame.pack_propagate(False)
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="📊 GÉNÉRATEUR DE RAPPORTS UGP",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(expand=True)
        
        # Section des fichiers
        self.create_files_section(main_container)
        
        # Section des paramètres
        self.create_params_section(main_container)
        
        # Boutons d'action
        self.create_action_buttons(main_container)
        
        # Barre de progression
        self.create_progress_section(main_container)
        
        # Zone de logs
        self.create_log_section(main_container)
    
    def create_files_section(self, parent):
        """Créer la section de sélection des fichiers"""
        files_frame = ctk.CTkFrame(parent)
        files_frame.pack(fill="x", pady=(0, 15))
        
        label = ctk.CTkLabel(
            files_frame,
            text="📁 FICHIERS REQUIS",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # Fichiers à sélectionner
        files_config = [
            ("BulkReport CSV:", "bulk", "*.csv"),
            ("Export Excel:", "export", "*.xlsx"),
            ("Table des Frais:", "fees", "*.xlsx"),
            ("Template Rapport:", "template", "*.xlsx")
        ]
        
        for label_text, key, file_type in files_config:
            self.create_file_row(files_frame, label_text, key, file_type)
    
    def create_file_row(self, parent, label_text, key, file_type):
        """Créer une ligne de sélection de fichier"""
        row_frame = ctk.CTkFrame(parent, fg_color="transparent")
        row_frame.pack(fill="x", padx=10, pady=5)
        
        # Label
        label = ctk.CTkLabel(row_frame, text=label_text, width=120, anchor="w")
        label.pack(side="left", padx=(10, 5))
        
        # Entry
        entry = ctk.CTkEntry(
            row_frame,
            textvariable=self.file_paths[key],
            width=400,
            placeholder_text=f"Sélectionner un fichier {file_type}"
        )
        entry.pack(side="left", padx=5)
        
        # Bouton parcourir pour tous les fichiers
        btn = ctk.CTkButton(
            row_frame,
            text="Parcourir",
            width=100,
            command=lambda k=key, ft=file_type: self.browse_file(k, ft)
        )
        btn.pack(side="left", padx=5)
        
        # Indicateur de statut
        self.create_status_indicator(row_frame, key)
    
    def create_status_indicator(self, parent, key):
        """Créer un indicateur de statut pour un fichier"""
        indicator = ctk.CTkLabel(
            parent,
            text="",
            width=30,
            font=ctk.CTkFont(size=16)
        )
        indicator.pack(side="left", padx=5)
        setattr(self, f"status_{key}", indicator)
    
    def create_params_section(self, parent):
        """Créer la section des paramètres"""
        params_frame = ctk.CTkFrame(parent)
        params_frame.pack(fill="x", pady=(0, 15))
        
        label = ctk.CTkLabel(
            params_frame,
            text="⚙️ PARAMÈTRES DU RAPPORT",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # Grille de paramètres
        params_grid = ctk.CTkFrame(params_frame, fg_color="transparent")
        params_grid.pack(fill="x", padx=20, pady=10)
        
        # Date
        self.create_param_field(params_grid, "Date:", self.metadata['date_paiement'], 0, 0)
        
        # Libellé
        self.create_param_field(params_grid, "Libellé:", self.metadata['libelle'], 0, 2, width=300)
        
        # Budget
        self.create_param_field(params_grid, "Budget (FCFA):", self.metadata['budget'], 1, 0)
        
        # Projet
        self.create_param_field(params_grid, "Projet:", self.metadata['projet'], 1, 2)
    
    def create_param_field(self, parent, label_text, variable, row, col, width=200):
        """Créer un champ de paramètre"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=col, columnspan=2 if width > 200 else 1, 
                  padx=10, pady=5, sticky="ew")
        
        label = ctk.CTkLabel(frame, text=label_text, width=100, anchor="w")
        label.pack(side="left", padx=(0, 5))
        
        entry = ctk.CTkEntry(frame, textvariable=variable, width=width)
        entry.pack(side="left", fill="x", expand=True)
    
    def create_action_buttons(self, parent):
        """Créer les boutons d'action"""
        buttons_frame = ctk.CTkFrame(parent, fg_color="transparent")
        buttons_frame.pack(fill="x", pady=10)
        
        # Bouton Générer
        self.generate_btn = ctk.CTkButton(
            buttons_frame,
            text="▶️ GÉNÉRER RAPPORT",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=40,
            width=200,
            fg_color="#28a745",
            hover_color="#218838",
            command=self.generate_report
        )
        self.generate_btn.pack(side="left", padx=10)
        
        # Bouton Ouvrir Dossier
        open_folder_btn = ctk.CTkButton(
            buttons_frame,
            text="📁 Ouvrir Dossier",
            height=40,
            width=150,
            command=self.open_output_folder
        )
        open_folder_btn.pack(side="left", padx=5)
        
        # Bouton Effacer
        clear_btn = ctk.CTkButton(
            buttons_frame,
            text="🔄 Réinitialiser",
            height=40,
            width=150,
            fg_color="#6c757d",
            hover_color="#545b62",
            command=self.clear_form
        )
        clear_btn.pack(side="left", padx=5)
    
    def create_progress_section(self, parent):
        """Créer la section de progression"""
        self.progress_frame = ctk.CTkFrame(parent)
        self.progress_frame.pack(fill="x", pady=10)
        self.progress_frame.pack_forget()  # Caché par défaut
        
        self.progress_label = ctk.CTkLabel(
            self.progress_frame,
            text="Traitement en cours...",
            font=ctk.CTkFont(size=12)
        )
        self.progress_label.pack(pady=(5, 0))
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, width=500)
        self.progress_bar.pack(pady=5)
        self.progress_bar.set(0)
    
    def create_log_section(self, parent):
        """Créer la zone de logs"""
        log_frame = ctk.CTkFrame(parent)
        log_frame.pack(fill="both", expand=True, pady=(10, 0))
        
        label = ctk.CTkLabel(
            log_frame,
            text="📋 JOURNAL D'ACTIVITÉ",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(anchor="w", padx=10, pady=5)
        
        # Textbox pour les logs
        self.log_text = ctk.CTkTextbox(
            log_frame,
            height=150,
            font=ctk.CTkFont(family="Consolas", size=11)
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    
    def browse_file(self, key, file_type):
        """Ouvrir le dialogue de sélection de fichier"""
        if file_type == "*.csv":
            filetypes = [("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        else:
            filetypes = [("Fichiers Excel", "*.xlsx;*.xls"), ("Tous les fichiers", "*.*")]
        
        # Utiliser le dossier UGP comme dossier initial
        initial_dir = r"C:\Users\faycalhabibahmat\Desktop\Moov\UGP"
        if not os.path.exists(initial_dir):
            initial_dir = os.getcwd()
        
        filename = filedialog.askopenfilename(
            title=f"Sélectionner le fichier {key}",
            initialdir=initial_dir,
            filetypes=filetypes
        )
        
        if filename:
            self.file_paths[key].set(filename)
            self.update_status(key, "✅")
            self.log(f"✅ Fichier sélectionné: {os.path.basename(filename)}")
    
    def update_status(self, key, status):
        """Mettre à jour l'indicateur de statut"""
        indicator = getattr(self, f"status_{key}", None)
        if indicator:
            indicator.configure(text=status)
    
    def log(self, message):
        """Ajouter un message au journal"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.update()
    
    def validate_inputs(self):
        """Valider les entrées avant traitement"""
        # Vérifier les fichiers requis
        required_files = ['bulk', 'export', 'fees']
        for key in required_files:
            file_path = self.file_paths[key].get()
            if not file_path or not os.path.exists(file_path):
                messagebox.showerror(
                    "Fichier manquant",
                    f"Veuillez sélectionner le fichier {key}"
                )
                return False
        
        # Vérifier le budget
        try:
            budget = int(self.metadata['budget'].get().replace(' ', ''))
            if budget <= 0:
                raise ValueError
        except:
            messagebox.showerror(
                "Budget invalide",
                "Le budget doit être un nombre positif"
            )
            return False
        
        return True
    
    def generate_report(self):
        """Générer le rapport dans un thread séparé"""
        if self.processing:
            return
        
        if not self.validate_inputs():
            return
        
        self.processing = True
        self.generate_btn.configure(state="disabled", text="⏳ Traitement...")
        self.progress_frame.pack(fill="x", pady=10)
        
        # Lancer le traitement dans un thread
        thread = threading.Thread(target=self.process_report)
        thread.daemon = True
        thread.start()
    
    def process_report(self):
        """Processus de génération du rapport"""
        try:
            # Initialiser les modules
            file_handler = FileHandler()
            processor = DataProcessor()
            generator = ReportGenerator()
            
            # Étape 1: Chargement des fichiers
            self.update_progress(0.2, "Chargement des fichiers...")
            
            bulk_df, bulk_metadata = file_handler.read_bulk_report(
                self.file_paths['bulk'].get()
            )
            self.log(f"✅ {len(bulk_df)} transactions chargées")
            
            export_df = file_handler.read_export_file(
                self.file_paths['export'].get()
            )
            self.log(f"✅ {len(export_df)} bénéficiaires chargés")
            
            fees_df = file_handler.read_fees_file(
                self.file_paths['fees'].get()
            )
            self.log(f"✅ Table des frais chargée")
            
            # Étape 2: Traitement des données
            self.update_progress(0.5, "Traitement des données...")
            
            # Préparer les métadonnées
            metadata = {
                'date_paiement': self.metadata['date_paiement'].get(),
                'libelle': self.metadata['libelle'].get(),
                'budget': int(self.metadata['budget'].get().replace(' ', '')),
                'projet': self.metadata['projet'].get(),
                'plan_name': bulk_metadata.get('plan_name', ''),
                'organization': bulk_metadata.get('organization', '')
            }
            
            # Traiter les transactions
            processed_df, errors = processor.process_transactions(
                bulk_df, export_df, fees_df, metadata
            )
            
            # Afficher les erreurs/warnings
            for error in errors:
                self.log(error)
            
            # Étape 3: Génération du rapport
            self.update_progress(0.8, "Génération du rapport Excel...")
            
            # Générer le nom du fichier
            date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"Rapport_{metadata['projet']}_{date_str}.xlsx"
            
            # Générer le rapport
            output_path = generator.generate_report(
                processed_df,
                metadata,
                output_name
            )
            
            # Obtenir les statistiques
            stats = processor.get_summary_stats(processed_df)
            
            # Succès!
            self.update_progress(1.0, "Terminé!")
            self.log(f"✅ Rapport généré avec succès: {output_name}")
            self.log(f"📊 {stats['total_transactions']} transactions, Total: {stats['total_amount']:,.0f} FCFA")
            
            # Ouvrir le fichier si configuré
            if self.ask_open_file():
                self.open_file(output_path)
            
        except Exception as e:
            self.log(f"❌ Erreur: {str(e)}")
            logger.error(f"Erreur génération rapport: {traceback.format_exc()}")
            messagebox.showerror("Erreur", f"Erreur lors de la génération:\n{str(e)}")
        
        finally:
            # Réinitialiser l'interface
            self.processing = False
            self.generate_btn.configure(state="normal", text="▶️ GÉNÉRER RAPPORT")
            self.progress_frame.pack_forget()
    
    def update_progress(self, value, message):
        """Mettre à jour la barre de progression"""
        self.progress_bar.set(value)
        self.progress_label.configure(text=message)
        self.update()
    
    def ask_open_file(self):
        """Demander si l'utilisateur veut ouvrir le fichier"""
        return messagebox.askyesno(
            "Rapport généré",
            "Le rapport a été généré avec succès.\nVoulez-vous l'ouvrir maintenant?"
        )
    
    def open_file(self, file_path):
        """Ouvrir un fichier avec l'application par défaut"""
        try:
            if platform.system() == 'Windows':
                os.startfile(file_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', file_path])
            else:  # Linux
                subprocess.call(['xdg-open', file_path])
        except Exception as e:
            self.log(f"⚠ Impossible d'ouvrir le fichier: {e}")
    
    def open_output_folder(self):
        """Ouvrir le dossier de sortie"""
        output_dir = os.path.abspath("outputs")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        try:
            if platform.system() == 'Windows':
                os.startfile(output_dir)
            elif platform.system() == 'Darwin':
                subprocess.call(['open', output_dir])
            else:
                subprocess.call(['xdg-open', output_dir])
            self.log(f"📁 Dossier ouvert: {output_dir}")
        except Exception as e:
            self.log(f"⚠ Impossible d'ouvrir le dossier: {e}")
    
    def clear_form(self):
        """Réinitialiser le formulaire"""
        for key in self.file_paths:
            if key != 'template':
                self.file_paths[key].set("")
                self.update_status(key, "")
        
        self.log_text.delete("1.0", "end")
        self.log("🔄 Formulaire réinitialisé")


def main():
    """Point d'entrée principal"""
    try:
        app = UGPReporterApp()
        app.mainloop()
    except Exception as e:
        logger.error(f"Erreur fatale: {traceback.format_exc()}")
        messagebox.showerror("Erreur fatale", f"L'application a rencontré une erreur:\n{str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
