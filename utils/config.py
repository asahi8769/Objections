import os

PYTHON = r"C:\Users\glovis-laptop\AppData\Local\Programs\Python\Python37"
VENV_64_DIR = os.path.join(os.getcwd(), 'venv')
SCRIPTS_DIR = os.path.join(VENV_64_DIR, 'Scripts')

URL = "https://partner.hyundai.com/gscm/"
GSCM_ID = os.environ.get ('GSCM_ID').upper()
GSCM_PW = os.environ.get ('GSCM_PW')
COORDINATES = {'QAMENU' : (309, 178)}


REPOSITORY = r'https://github.com/asahi8769/Objections.git'


