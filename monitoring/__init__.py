"""
Module de monitoring automatique pour UGP Reporter
"""

from .file_watcher import SmartFileWatcher
from .pdf_converter import ProfessionalPDFConverter
from .email_sender import ProfessionalEmailSender
from .auto_processor import AutoProcessor

__all__ = [
    'SmartFileWatcher',
    'ProfessionalPDFConverter',
    'ProfessionalEmailSender',
    'AutoProcessor'
]

__version__ = '1.0.0'
