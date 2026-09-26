"""Foglio Rate, colonna Scaduto e comando `situazione` (0.4.0)."""
import json
import unittest

import openpyxl

from helpers import CondominioTest

SCADENZE = ["2026-01-31", "2026-04-30", "2026-07-31", "2026-10-31"]


class TestRate(CondominioTest):
    def setUp(self):
        super().setUp()
        for i, d in enumerate(SCADENZE, start=1):  # U03: 4 rate da 100 €
            self.rata("U03", i, d, 100)

    def rata(self, uid, n, scadenza, importo, prospetto="2026-01-10_riparto_preventivo-2026.xlsx", **extra):
        riga = {"ID unità": uid, "Esercizio": 2026, "Prospetto": prospetto, "Rata": n, "Scadenza": scadenza, "Importo": importo, **extra}
        code, out = self.ris("append", "Rate", json.dumps(riga))
        self.assertEqual(code, 0, out)

    def versa(self, importo, data="2026-02-01", uid="U03"):
        self.ris("append", "Versamenti", json.dumps({"Data": data, "ID unità": uid, "Importo": importo, "Esercizio": 2026}))

    def sit(self, oggi="2026-09-26", uid="U03"):
        code, out = self.ris("--oggi", oggi, "situazione", "--unita", uid)
        self.assertEqual(code, 0, out)
        self.assertEqual(len(out["unita"]), 1)
        return out["unita"][0]

    def stati(self, s):
        return [(r["Stato"], r["Pagato"]) for r in s["rate"]]

    def test_dovuto_dalle_rate_e_nessun_pagamento(self):
        s = self.sit()
        self.assertEqual((s["Dovuto"], s["Versato"], s["Saldo"], s["Scaduto"]), (400, 0, 400, 300))
        self.assertEqual(self.stati(s), [("scaduta", 0), ("scaduta", 0), ("scaduta", 0), ("da pagare", 0)])

    def test_pagamento_parziale_imputato_in_ordine_di_scadenza(self):
        self.versa(150)
        s = self.sit()
        self.assertEqual((s["Versato"], s["Saldo"], s["Scaduto"]), (150, 250, 150))
        self.assertEqual(self.stati(s), [("pagata", 100), ("scaduta", 50), ("scaduta", 0), ("da pagare", 0)])

    def test_rata_non_scaduta_coperta_in_parte(self):
        self.versa(350)
        s = self.sit()
        self.assertEqual(s["Scaduto"], 0)
        self.assertEqual(self.stati(s)[3], ("parziale", 50))

    def test_il_giorno_della_scadenza_non_e_ritardo(self):
        self.versa(200)
        s = self.sit(oggi="2026-07-31")
        self.assertEqual(s["Scaduto"], 0)
        self.assertEqual(s["rate"][2]["Stato"], "da pagare")
        self.assertEqual(self.sit(oggi="2026-08-01")["Scaduto"], 100)

    def test_valori_scritti_nel_foglio_situazione(self):
        self.versa(150)
        self.ris("--oggi", "2026-09-26", "verifica")
        riga = self.situazione("U03")
        self.assertEqual((riga["Dovuto"], riga["Versato"], riga["Saldo"], riga["Scaduto"]), (400, 150, 250, 150))

    def test_rate_sostituite_non_contano(self):
        self.ris("update", "Rate", "--where", "ID unità=U03", '{"Valida": "NO"}')
        self.rata("U03", 1, "2026-11-30", 380, prospetto="2026-11-01_riparto_consuntivo-2026.xlsx")
        s = self.sit()
        self.assertEqual((s["Dovuto"], s["Scaduto"]), (380, 0))
        self.assertEqual(len(s["rate"]), 1)

    def test_tutte_le_rate_sostituite_azzera_il_dovuto(self):
        self.ris("update", "Situazione", "--where", "ID unità=U03", '{"Dovuto": 999}')
        self.ris("update", "Rate", "--where", "ID unità=U03", '{"Valida": "NO"}')
        self.assertEqual(self.sit()["Dovuto"], 0)

    def test_rate_di_altro_esercizio_ignorate(self):
        self.ris("append", "Rate", json.dumps({"ID unità": "U03", "Esercizio": 2025, "Rata": 1, "Scadenza": "2025-06-30", "Importo": 500}))
        self.assertEqual(self.sit()["Dovuto"], 400)

    def test_versamenti_elencati(self):
        self.versa(150, data="2026-03-05")
        v = self.sit()["versamenti"]
        self.assertEqual([(x["Data"], x["Importo"]) for x in v], [("2026-03-05", 150)])


class TestSenzaRate(CondominioTest):
    """Registri 0.3: nessuna riga in Rate, il Dovuto scritto resta e lo Scaduto è ignoto (vuoto)."""

    def test_dovuto_scritto_e_scaduto_vuoto(self):
        self.ris("update", "Situazione", "--where", "ID unità=U01", '{"Dovuto": 416.25}')
        _, out = self.ris("situazione", "--unita", "U01")
        s = out["unita"][0]
        self.assertEqual((s["Dovuto"], s["Versato"], s["Saldo"], s["Scaduto"], s["rate"]), (416.25, 150, 266.25, None, []))
        self.assertIsNone(self.situazione("U01")["Scaduto"])

    def test_tutte_le_unita_e_filtri(self):
        _, out = self.ris("situazione")
        self.assertEqual([u["ID unità"] for u in out["unita"]], ["U01", "U02", "U03", "U04"])
        _, out = self.ris("situazione", "--esercizio", 2025)
        self.assertEqual(out["unita"], [])

    def test_errori_leggibili(self):
        code, out = self.ris("situazione", "--unita", "U99")
        self.assertNotEqual(code, 0)
        self.assertIn("U99", out)
        code, out = self.reg("situazione")
        self.assertNotEqual(code, 0)
        self.assertIn("registro-riservato.xlsx", out)


class TestAggiornamentoDa03(CondominioTest):
    def test_schema_aggiunge_rate_scaduto_e_id_movimento(self):
        path = self.path("registro-riservato.xlsx")
        wb = openpyxl.load_workbook(path)
        del wb["Rate"]
        wb["Situazione"].delete_cols(10)  # Scaduto
        wb["Versamenti"].delete_cols(9)   # ID movimento
        wb.save(path)
        _, prima = self.ris("read", "Versamenti")
        code, out = self.ris("schema")
        self.assertEqual(code, 0)
        self.assertEqual(sorted(out["aggiunte"]), sorted([
            "Situazione: colonna 'Scaduto' (in coda)",
            "Versamenti: colonna 'ID movimento' (in coda)",
            "foglio 'Rate'",
        ]))
        _, dopo = self.ris("read", "Versamenti")
        self.assertEqual([{k: v for k, v in d.items() if k != "ID movimento"} for d in dopo], prima)
        code, out = self.ris("situazione", "--unita", "U01")
        self.assertEqual((code, out["unita"][0]["Versato"]), (0, 150))


if __name__ == "__main__":
    unittest.main()
