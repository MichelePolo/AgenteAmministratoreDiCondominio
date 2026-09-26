"""Istruzioni dell'automazione (0.5.0): un compito minimo, senza percorsi e senza dati per unità."""
import os
import re
import unittest

from helpers import SKILLS

ISTRUZIONI = os.path.join(SKILLS, "automazioni", "references", "istruzioni-riepilogo.md")


class TestIstruzioni(unittest.TestCase):
    def setUp(self):
        with open(ISTRUZIONI, encoding="utf-8") as f:
            self.testo = f.read()

    def test_segnaposti(self):
        self.assertEqual(sorted(set(re.findall(r"\{(\w+)\}", self.testo))), ["email", "nome"])
        compilato = self.testo.format(nome="Condominio Prova", email="admin@example.it")
        self.assertNotIn("{", compilato)

    def test_solo_script_di_stato_e_invio_gmail(self):
        self.assertIn("amministratore-condominio:stato", self.testo)
        self.assertIn("--testo", self.testo)
        self.assertIn("Gmail", self.testo)
        skill = set(re.findall(r"amministratore-condominio:([\w-]+)", self.testo))
        self.assertEqual(skill, {"stato"})
        for vietato in ("registro.py", "riparto", "abbina", "situazione_personale", "archivia", "sollecit"):
            self.assertNotIn(vietato, self.testo)

    def test_nessun_percorso(self):
        """Il plugin installato cambia cartella a ogni versione: le istruzioni non citano percorsi."""
        self.assertNotRegex(self.testo, r"(/home/|~/|\.claude/|/cache/|\\\\|[A-Z]:\\\\|\.py\b)")

    def test_un_solo_destinatario(self):
        self.assertEqual(self.testo.count("{email}"), 2)  # destinatario e caso di errore
        self.assertIn("nessun Cc, nessun Bcc", self.testo)


if __name__ == "__main__":
    unittest.main()
