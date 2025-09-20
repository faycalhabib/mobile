"""
Script pour basculer facilement entre le mode classique et le mode optimisé
"""
import json
from pathlib import Path

def toggle_fast_mode():
    """Bascule entre le mode classique et le mode rapide"""
    
    config_path = Path('config/auto_processor_config.json')
    
    # Lire la config actuelle
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    # Obtenir l'état actuel
    current_state = config.get('optimization', {}).get('use_fast_mode', False)
    
    # Basculer
    new_state = not current_state
    
    # S'assurer que la section optimization existe
    if 'optimization' not in config:
        config['optimization'] = {}
    
    config['optimization']['use_fast_mode'] = new_state
    
    # Sauvegarder
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)
    
    # Afficher le résultat
    print("=" * 50)
    print(" CONFIGURATION DU MODE D'ÉCRITURE EXCEL")
    print("=" * 50)
    print()
    
    if new_state:
        print("🚀 MODE RAPIDE ACTIVÉ (openpyxl)")
        print()
        print("Avantages:")
        print("  ✓ 5-10x plus rapide")
        print("  ✓ Utilise moins de mémoire")
        print("  ✓ Ne nécessite pas Excel ouvert")
        print()
        print("Note: Fallback automatique au mode classique si erreur")
    else:
        print("🔧 MODE CLASSIQUE ACTIVÉ (Win32COM)")
        print()
        print("Avantages:")
        print("  ✓ Format 100% identique")
        print("  ✓ Compatible avec toutes les fonctionnalités Excel")
        print("  ✓ Mode le plus stable")
    
    print()
    print("Pour changer à nouveau, relancez ce script.")
    print("=" * 50)

def show_current_mode():
    """Affiche le mode actuellement configuré"""
    
    config_path = Path('config/auto_processor_config.json')
    
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    current_state = config.get('optimization', {}).get('use_fast_mode', False)
    
    print()
    print(f"Mode actuel: {'🚀 RAPIDE (openpyxl)' if current_state else '🔧 CLASSIQUE (Win32COM)'}")
    print()

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == 'status':
        show_current_mode()
    else:
        toggle_fast_mode()
