"""Comportamento di riparto.py (versione 0.3): somma esatta al centesimo, rate, conduttori, errori."""
import os
import unittest
from decimal import Decimal

from helpers import CondominioTest


def cent(x):
    return Decimal(str(x)).quantize(Decimal("0.01"))


class TestRiparto(CondominioTest):
    def test_somma_esatta_e_rate(self):
        code, out = self.riparto("--rate", 4)
        self.assertEqual(code, 0, out)
        self.assertEqual((out["spese"], out["totale"]), (3, 2442.5))
        per_unita = out["per_unita"]
        self.assertEqual(sum(cent(u["totale"]) for u in per_unita.values()), cent(2442.5))
        for u in per_unita.values():
            self.assertEqual(len(u["rate"]), 4)
            self.assertEqual(sum(cent(r) for r in u["rate"]), cent(u["totale"]))
        self.assertEqual(per_unita["U01"]["totale"], 416.25)

    def test_conduttore_solo_per_unita_affittata(self):
        _, out = self.riparto()
        self.assertEqual(out["unita_con_conduttore"], 1)
        for uid, u in out["per_unita"].items():
            self.assertEqual(cent(u["a_carico_proprieta"]) + cent(u["a_carico_conduttore"]), cent(u["totale"]))
            if uid != "U02":
                self.assertEqual(u["a_carico_conduttore"], 0)
        # U02: ENEL (B 250‰ di 412,50 = 103,125 → 103,12: il centesimo va a un'unità con resto maggiore)
        # + Otis (C 300‰ di 780 = 234), entrambe al 100%; la polizza (0%) resta al proprietario
        self.assertEqual(cent(out["per_unita"]["U02"]["a_carico_conduttore"]), cent("103.12") + cent(234))

    def test_prospetto_mai_sovrascritto(self):
        _, a = self.riparto("--titolo", "Prova")
        _, b = self.riparto("--titolo", "Prova")
        self.assertTrue(a["prospetto"].endswith("_bozza.xlsx"))
        self.assertEqual(b["prospetto"], a["prospetto"].replace(".xlsx", "_2.xlsx"))
        self.assertTrue(os.path.exists(self.path(a["prospetto"])))
        self.assertTrue(os.path.exists(self.path(b["prospetto"])))

    def test_nota_di_credito_ripartita_con_segno(self):
        self.reg("append", "Spese", '{"Data": "2026-06-01", "Fornitore": "ENEL", "Importo": -10, "Tabella": "A", "Esercizio": 2026}')
        _, out = self.riparto("--ids", "4")
        self.assertEqual(out["totale"], -10)
        self.assertEqual(sum(cent(u["totale"]) for u in out["per_unita"].values()), cent(-10))

    def test_spesa_a_carico_di_una_unita(self):
        self.reg("append", "Spese", '{"Data": "2026-06-01", "Fornitore": "Fabbro", "Importo": 80, "Tabella": "UNITA:U03", "Esercizio": 2026}')
        _, out = self.riparto("--ids", "4")
        self.assertEqual(out["per_unita"]["U03"]["totale"], 80)
        self.assertEqual(out["a_carico_singola_unita"], {"UNITA:U03": 80})

    def test_rifiuta_spesa_senza_tabella(self):
        self.reg("update", "Spese", "--where", "ID=1", '{"Tabella": ""}')
        code, out = self.riparto()
        self.assertEqual(code, 2)
        self.assertFalse(out["ok"])
        self.assertTrue(any("Spesa 1" in e and "manca la Tabella" in e for e in out["errori"]), out)

    def test_rifiuta_millesimi_che_non_quadrano(self):
        self.reg("update", "Millesimi", "--where", "ID unità=U01", '{"Tab A": 290}')
        code, out = self.riparto()
        self.assertEqual(code, 2)
        self.assertIn("Tab A somma a 990, non a 1000", out["errori"])


if __name__ == "__main__":
    unittest.main()
