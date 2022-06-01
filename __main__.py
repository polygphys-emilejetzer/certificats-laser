# -*- coding: utf-8 -*-
"""Créer de nouveaux certificats laser au besoin."""

# Bibliothèque standard
import subprocess
import time

from pathlib import Path

# Bibliothèque PIPy
import schedule

# Imports relatifs
from . import SSTLaserCertificatsConfig, SSTLaserCertificatsForm
from ...outils.reseau import OneDrive

chemin_config = Path('~').expanduser() / 'certificats_laser.cfg'
config = SSTLaserCertificatsConfig(chemin_config)

dossier = OneDrive('',
                   config.get('onedrive', 'organisation'),
                   config.get('onedrive', 'sous-dossier'),
                   partagé=True)
fichier = dossier / config.get('formulaire', 'nom')
config.set('formulaire', 'chemin', str(fichier))

formulaire = SSTLaserCertificatsForm(config)

exporteur = subprocess.Popen(['unoconv', '--listener'])

schedule.every().day.at('08:00').do(formulaire.mise_à_jour)

formulaire.mise_à_jour()
try:
    while True:
        schedule.run_pending()
        time.sleep(1)
finally:
    exporteur.terminate()
