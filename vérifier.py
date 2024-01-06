# -*- coding: utf-8 -*-

import time

from datetime import datetime, timedelta
from argparse import ArgumentParser
from pathlib import Path

import schedule

from polygphys.sst.certificats_laser.nouveau_certificats import SSTLaserCertificatsConfig, FormationLaser

def main():
    chemin_config = Path('~').expanduser() / 'certificats_laser.cfg'
    config = SSTLaserCertificatsConfig(chemin_config)

    adresse = config.get('bd', 'adresse')
    tableau = FormationLaser(adresse)

    parseur = ArgumentParser(description='Spécifier les critères de recherche.')
    parseur.add_argument('--matricule', type=str)
    parseur.add_argument('--nom', type=str)
    parseur.add_argument('--courriel', type=str)
    parseur.add_argument('--ans', type=int, default=3)
    arguments = parseur.parse_args()

    date = datetime.today() - timedelta(days=365*arguments.ans)
    rés = tableau.trouver(arguments.matricule,
                          arguments.nom,
                          arguments.courriel,
                          date)

    if not rés.empty:
        print(f'Dernière formation reçue le {rés.date.iloc[-1]}.')
    else:
        print(f'Aucune formation reçue depuis {date}.')
    

# Programme
if __name__ == '__main__':
        schedule.every().day.at('08:00').do(main)

        main()
