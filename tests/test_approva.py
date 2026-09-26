"""riparto.py approva (0.4.0): rate, Situazione, Scadenze, Diario, rinomina, idempotenza, sostituzioni."""
import os
import shutil
import unittest
from decimal import Decimal

from helpers import CondominioTest, RIPARTO, run

QUATTRO = "2026-01-31,2026-04-30,2026-07-31,2026-10-31"


def cent(x):
    return Decimal(str(x)).quantize(Decimal("0.01"))


class TestApprova(CondominioTest):
    def bozza(self, *args):
        code, out = self.riparto(*args)
        self.assertEqual(code, 0, out)
        return out

    def approva(self, prospetto, scadenze=QUATTRO, *extra):
        return run(RIPARTO, "approva", "--dir", self.dir, "--prospetto", prospetto, "--scadenze", scadenze,
                   "--approvato", "Michele", *extra)

    def rate(self):
        return self.ris("read", "Rate")[1]

    def test_approvazione_completa(self):
        prev = self.bozza("--rate", 4, "--titolo", "Preventivo 2026")
        code, out = self.approva(prev["prospetto"], QUATTRO, "--verbale", "assemblea del 10 gennaio")
        self.assertEqual(code, 0, out)
        definitivo = prev["prospetto"].replace("_bozza.xlsx", ".xlsx")
        self.assertEqual(out["prospetto"], definitivo)
        self.assertTrue(os.path.exists(self.path(definitivo)))
        self.assertFalse(os.path.exists(self.path(prev["prospetto"])))
        # 4 unità × 4 rate, importi identici al prospetto
        rate = self.rate()
        self.assertEqual(len(rate), 16)
        self.assertTrue(all(r["Prospetto"] == os.path.basename(definitivo) and r["Valida"] == "SI" for r in rate))
        for uid, u in prev["per_unita"].items():
            mie = sorted((r for r in rate if r["ID unità"] == uid), key=lambda r: r["Rata"])
            self.assertEqual([cent(r["Importo"]) for r in mie], [cent(x) for x in u["rate"]])
            self.assertEqual([r["Scadenza"] for r in mie], QUATTRO.split(","))
            self.assertEqual(cent(self.situazione(uid)["Dovuto"]), cent(u["totale"]))
        self.assertEqual(sum(cent(self.situazione(u)["Dovuto"]) for u in prev["per_unita"]), cent(2442.5))
        # Scadenze pubbliche: solo date, nessun nome né importo per unità
        _, scad = self.reg("read", "Scadenze", "--where", "Tipo=rata")
        self.assertEqual([s["Data"] for s in scad], QUATTRO.split(","))
        self.assertEqual(scad[1]["Descrizione"], "Rata 2 di 4 — Preventivo 2026")
        testo = str(scad)
        for u in prev["per_unita"].values():
            self.assertNotIn(u["intestatario"], testo)
        # Diario
        _, diario = self.reg("read", "Diario")
        self.assertEqual(diario[-1]["Operazione"], "Approvato prospetto di riparto")
        self.assertEqual(diario[-1]["Approvato da"], "Michele")
        self.assertIn("assemblea del 10 gennaio", diario[-1]["Dettaglio"])

    def test_seconda_approvazione_rifiutata(self):
        prev = self.bozza("--rate", 4)
        self.approva(prev["prospetto"])
        code, out = self.approva(prev["prospetto"])
        self.assertEqual(code, 2)
        self.assertIn("già approvato?", out["errori"][0])
        self.assertEqual(len(self.rate()), 16)

    def test_errori_prima_di_qualsiasi_scrittura(self):
        prev = self.bozza("--rate", 4)
        code, out = self.approva(prev["prospetto"], "2026-01-31,2026-04-30")
        self.assertEqual(code, 2)
        self.assertIn("Il prospetto ha 4 rate, sono state indicate 2 scadenze", out["errori"])
        code, out = self.approva(prev["prospetto"], "2026-04-30,2026-01-31,2026-07-31,2026-10-31")
        self.assertIn("Le scadenze vanno indicate in ordine di data", out["errori"])
        code, out = self.approva(prev["prospetto"], "31/01/2026")
        self.assertIn("Scadenze non valide", out["errori"][0])
        self.assertEqual(self.rate(), [])
        self.assertTrue(os.path.exists(self.path(prev["prospetto"])))
        _, scad = self.reg("read", "Scadenze")
        self.assertEqual(scad, [])

    def test_prospetto_senza_rate_una_scadenza(self):
        prev = self.bozza("--ids", "3", "--titolo", "Polizza")
        code, out = self.approva(prev["prospetto"], "2026-11-30")
        self.assertEqual(code, 0, out)
        self.assertEqual(len(self.rate()), 4)
        self.assertEqual(self.situazione("U01")["Dovuto"], 375)  # 1250 × 300/1000
        _, scad = self.reg("read", "Scadenze")
        self.assertEqual(scad[0]["Descrizione"], "Pagamento — Polizza")

    def test_consuntivo_sostituisce_preventivo(self):
        prev = self.bozza("--rate", 4, "--titolo", "Preventivo 2026")
        _, a = self.approva(prev["prospetto"])
        cons = self.bozza("--ids", "1,2", "--rate", 1, "--titolo", "Consuntivo 2026")
        code, out = self.approva(cons["prospetto"], "2026-12-15", "--sostituisce", a["prospetto"])
        self.assertEqual(code, 0, out)
        self.assertEqual(out["rate_sostituite"], 16)
        vecchie = [r for r in self.rate() if r["Prospetto"] == os.path.basename(a["prospetto"])]
        self.assertTrue(all(r["Valida"] == "NO" and "sostituita da" in r["Note"] for r in vecchie))
        for uid, u in cons["per_unita"].items():
            self.assertEqual(cent(self.situazione(uid)["Dovuto"]), cent(u["totale"]))

    def test_sostituisce_un_prospetto_mai_approvato(self):
        prev = self.bozza("--rate", 4)
        code, out = self.approva(prev["prospetto"], QUATTRO, "--sostituisce", "prospetti/inventato.xlsx")
        self.assertEqual(code, 2)
        self.assertIn("nessuna rata di quel prospetto", out["errori"][0])

    def test_riparto_rifatto_lo_stesso_giorno(self):
        prima = self.bozza("--rate", 4, "--titolo", "Prova")
        _, a = self.approva(prima["prospetto"])
        dopo = self.bozza("--rate", 4, "--titolo", "Prova")
        self.assertEqual(dopo["prospetto"], prima["prospetto"])  # la bozza è stata rinominata: il nome è di nuovo libero
        code, b = self.approva(dopo["prospetto"], QUATTRO, "--sostituisce", a["prospetto"])
        self.assertEqual(code, 0, b)
        self.assertEqual(b["prospetto"], a["prospetto"].replace(".xlsx", "_2.xlsx"))

    def test_dovuto_scritto_in_0_3_non_si_perde(self):
        self.ris("update", "Situazione", "--where", "ID unità=U01", '{"Dovuto": 500}')
        straord = self.bozza("--ids", "3", "--titolo", "Straordinario")
        code, out = self.approva(straord["prospetto"], "2026-11-30")
        self.assertEqual(code, 0, out)
        self.assertEqual(out["dovuti_precedenti_riportati"], {"U01": 500})
        self.assertEqual(self.situazione("U01")["Dovuto"], 875)
        _, sit = self.ris("--oggi", "2026-12-15", "situazione", "--unita", "U01")
        rate = sit["unita"][0]["rate"]
        self.assertEqual((rate[0]["Rata"], rate[0]["Scadenza"], rate[0]["Pagato"]), (0, None, 150))  # imputata per prima
        self.assertEqual(sit["unita"][0]["Scaduto"], 375)  # la rata 0 non risulta mai scaduta

    def test_crea_le_righe_mancanti_di_situazione(self):
        prev = self.bozza("--rate", 2)
        code, out = self.approva(prev["prospetto"], "2027-03-31,2027-09-30", "--esercizio", 2027)
        self.assertEqual(code, 0, out)
        self.assertEqual(out["situazioni_create"], ["U01", "U02", "U03", "U04"])
        _, righe = self.ris("read", "Situazione", "--where", "Esercizio=2027")
        self.assertEqual(len(righe), 4)

    def test_registro_riservato_in_cartella_privata(self):
        os.makedirs(self.path("privato"))
        shutil.move(self.path("registro-riservato.xlsx"), self.path("privato", "registro-riservato.xlsx"))
        self.reg("set", "Condominio", "Percorso registro riservato", "privato")
        prev = self.bozza("--rate", 4)
        code, out = self.approva(prev["prospetto"])
        self.assertEqual(code, 0, out)
        _, rate = run(RIPARTO.replace("riparto/scripts/riparto.py", "condominio-base/scripts/registro.py"),
                      "--dir", self.path("privato"), "--file", "registro-riservato.xlsx", "read", "Rate")
        self.assertEqual(len(rate), 16)


if __name__ == "__main__":
    unittest.main()
