"""Aggiornamento di un registro 0.5 alla 0.6: foglio Ordinanti e colonna Spese.ID movimento."""
import unittest

import openpyxl

from helpers import CondominioTest


class TestAggiornamentoDa05(CondominioTest):
    def test_schema_aggiunge_ordinanti_e_id_movimento(self):
        rc, rr = self.path("registro-condominio.xlsx"), self.path("registro-riservato.xlsx")
        wb = openpyxl.load_workbook(rc)
        wb["Spese"].delete_cols(13)  # ID movimento
        wb.save(rc)
        wb = openpyxl.load_workbook(rr)
        del wb["Ordinanti"]
        wb.save(rr)
        _, prima = self.reg("read", "Spese")
        code, out = self.reg("schema")
        self.assertEqual((code, out["aggiunte"]), (0, ["Spese: colonna 'ID movimento' (in coda)"]))
        code, out = self.ris("schema")
        self.assertEqual((code, out["aggiunte"]), (0, ["foglio 'Ordinanti'"]))
        _, dopo = self.reg("read", "Spese")
        self.assertEqual([{k: v for k, v in d.items() if k != "ID movimento"} for d in dopo], prima)
        _, ordinanti = self.ris("read", "Ordinanti")
        self.assertEqual(ordinanti, [])


if __name__ == "__main__":
    unittest.main()
