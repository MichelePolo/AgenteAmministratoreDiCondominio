"""Invarianti sui file che i condomini aprono da Drive, e sulla scansione dell'inbox."""
import glob
import os
import shutil
import unittest

import openpyxl

from helpers import CondominioTest, SCANSIONA, run


class TestNessunaFormula(CondominioTest):
    """L'anteprima di Google Drive su cellulare mostra solo i valori salvati: niente formule, mai."""

    def test_registri_e_prospetti(self):
        self.reg("append", "Spese", '{"Data": "2026-06-01", "Fornitore": "X", "Importo": 99.99, "Tabella": "B", "Esercizio": 2026}')
        self.ris("append", "Versamenti", '{"Data": "2026-02-01", "ID unità": "U01", "Importo": 100}')
        self.reg("verifica")
        self.riparto("--rate", 3)
        files = glob.glob(self.path("*.xlsx")) + glob.glob(self.path("prospetti", "*.xlsx"))
        self.assertGreaterEqual(len(files), 3)
        for f in files:
            wb = openpyxl.load_workbook(f)
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        self.assertFalse(isinstance(c.value, str) and c.value.startswith("="),
                                         f"formula in {os.path.basename(f)}!{ws.title}!{c.coordinate}: {c.value}")


class TestScansiona(CondominioTest):
    def scansiona(self):
        code, out = run(SCANSIONA, "--dir", self.dir)
        self.assertEqual(code, 0, out)
        return {f["nome"]: f for f in out["file"]}, out

    def test_nuovi_duplicati_e_file_ignorati(self):
        inbox = self.path("da analizzare")
        with open(os.path.join(inbox, "bolletta.pdf"), "wb") as f:
            f.write(b"%PDF bolletta")
        shutil.copy(os.path.join(inbox, "bolletta.pdf"), os.path.join(inbox, "bolletta (1).pdf"))
        for ignorato in ("scorciatoia.gdoc", ".nascosto", "~$aperto.xlsx"):
            open(os.path.join(inbox, ignorato), "w").close()
        files, out = self.scansiona()
        self.assertEqual((out["totale"], out["nuovi"]), (2, 1))
        self.assertEqual(files["bolletta (1).pdf"]["stato"], "nuovo")  # ordine alfabetico: "(" < "."
        self.assertEqual(files["bolletta.pdf"]["stato"], "duplicato_in_inbox")

    def test_duplicato_di_file_gia_archiviato(self):
        inbox = self.path("da analizzare")
        with open(os.path.join(inbox, "enel.pdf"), "wb") as f:
            f.write(b"%PDF enel")
        files, _ = self.scansiona()
        h = files["enel.pdf"]["hash"]
        self.reg("append", "Indice", f'{{"Hash": "{h}", "Nome originale": "enel.pdf", "Percorso": "archivio/2026/fatture/x.pdf"}}')
        files, _ = self.scansiona()
        self.assertEqual(files["enel.pdf"]["stato"], "duplicato")
        self.assertEqual(files["enel.pdf"]["copia_archiviata"], "archivio/2026/fatture/x.pdf")


if __name__ == "__main__":
    unittest.main()
