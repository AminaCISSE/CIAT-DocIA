import getpass
import os
import subprocess
import sys

cle = getpass.getpass("Colle ta cle API OpenAI (elle restera masquee) : ")

if not cle.strip():
    print("Aucune cle saisie. Relance le programme.")
    sys.exit(1)

environnement = os.environ.copy()
environnement["OPENAI_API_KEY"] = cle.strip()

subprocess.run(
    [sys.executable, "-m", "streamlit", "run", "app.py"],
    env=environnement
)