import os
import netrc
from time import time
import requests
import glob
from datetime import datetime
import time
import csv
import re
import sys

# Set .netrc host, server URL, and log filename prefix.
netrc_host = 'kc.kobotoolbox.org'  # This should match entry in .netrc file
server_url = 'https://kc.kobotoolbox.org/submission'
log_prefix = 'kc'

# Récupération des identifiants depuis le .netrc du dossier courant
info = netrc.netrc(".netrc")
username, account, password = info.authenticators(netrc_host)

def progress(count, total, status=''):
    """
    Affiche une barre de progression dans le terminal.
    """
    bar_len = 60
    filled_len = int(round(bar_len * count / float(total)))
    percents = int(round(100 * count / float(total), 0))
    bar = '=' * filled_len + '-' * (bar_len - filled_len)
    sys.stdout.write('[%s] %s%s ...%s\r' % (bar, percents, '%', status))
    sys.stdout.flush()

# Recherche tous les fichiers XML à soumettre
filelist = glob.glob("tempfiles/*.xml")
logdata = []
logfilename = datetime.now().strftime('postlog__%Y-%m-%d_%H-%M.csv')

total = len(filelist)
for i, filepath in enumerate(filelist, 1):
    with open(filepath, 'rb') as f:
        # Prépare le fichier pour l'envoi POST
        files = {'xml_submission_file': (os.path.basename(filepath), f, 'text/xml')}
        response = requests.post(server_url, files=files, auth=(username, password))

        d = {}
        d['res'] = response
         # Extrait l'UUID du nom du fichier (entre le dernier / et .xml)
        match = re.search(r'([^/\\]+)\.xml$', filepath)
        d['uuid'] = match.group(1) if match else ''
        logdata.append(d)
        progress(i, total, status='Posting...')

print("\nWriting log file...")

# Écriture du fichier de log CSV
with open(log_prefix + logfilename, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = ['code', 'response_text', 'date', 'uuid']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, dialect='excel')
    writer.writeheader()

    for ld in logdata:
        writer.writerow({
            'code':          ld['res'].status_code,
            'response_text': ld['res'].text,
            'date':          ld['res'].headers.get('Date', ''),
            'uuid':          ld['uuid']
        })
    print(log_prefix + logfilename + " Complete")