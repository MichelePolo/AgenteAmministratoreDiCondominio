"""Utilità comuni ai test: una cartella di condominio temporanea e l'esecuzione degli script come li usa Claude."""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILLS = os.path.join(ROOT, "plugins", "amministratore-condominio", "skills")
CREA = os.path.join(SKILLS, "condominio-setup", "scripts", "crea_registro.py")
REGISTRO = os.path.join(SKILLS, "condominio-base", "scripts", "registro.py")
RIPARTO = os.path.join(SKILLS, "riparto", "scripts", "riparto.py")
SCANSIONA = os.path.join(SKILLS, "archivio", "scripts", "scansiona.py")
ESERCIZIO = 2026


def run(script, *args, cwd=None):
    """Esegue uno script con lo stesso interprete dei test; restituisce (codice, json o testo)."""
    p = subprocess.run([sys.executable, script, *map(str, args)], capture_output=True, text=True, cwd=cwd)
    out = p.stdout.strip()
    try:
        out = json.loads(out)
    except ValueError:
        out = out or p.stderr.strip()
    return p.returncode, out


class CondominioTest(unittest.TestCase):
    """Ogni test parte da una cartella nuova creata da crea_registro.py (con i dati di esempio)."""
    esempio = True

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="condominio-test-")
        self.addCleanup(shutil.rmtree, self.dir, ignore_errors=True)
        args = ["--dir", self.dir, "--nome", "Condominio Prova", "--esercizio", ESERCIZIO]
        code, out = run(CREA, *args, *(["--esempio"] if self.esempio else []))
        self.assertEqual(code, 0, out)

    # --- scorciatoie
    def reg(self, *args):
        return run(REGISTRO, "--dir", self.dir, *args)

    def ris(self, *args):
        return run(REGISTRO, "--dir", self.dir, "--file", "registro-riservato.xlsx", *args)

    def riparto(self, *args):
        return run(RIPARTO, "--dir", self.dir, *args)

    def path(self, *parti):
        return os.path.join(self.dir, *parti)

    def wb(self, nome="registro-condominio.xlsx"):
        return openpyxl.load_workbook(self.path(nome))

    def situazione(self, uid):
        code, rows = self.ris("read", "Situazione", "--where", f"ID unità={uid}")
        self.assertEqual(code, 0, rows)
        self.assertEqual(len(rows), 1, rows)
        return rows[0]
