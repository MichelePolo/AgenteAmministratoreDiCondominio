"""situazione_personale.py (0.4.0): un testo per unità, solo i suoi dati, copie nella cartella riservata."""
import json
import os
import unittest

from helpers import CondominioTest, RIPARTO, SKILLS, run

SCRIPT = os.path.join(SKILLS, "comunicazioni", "scripts", "situazione_personale.py")
OGGI = "2026-09-26"


class TestSituazionePersonale(CondominioTest):
    def setUp(self):
        super().setUp()
        self.reg("set", "Condominio", "Conto corrente", "IT60X0542811101000000123456")
        self.reg("set", "Condominio", "Amministratore", "Michele Polo")
        _, prev = self.riparto("--rate", 4, "--titolo", "Preventivo 2026")
        self.prev = prev["per_unita"]
        code, out = run(RIPARTO, "approva", "--dir", self.dir, "--prospetto", prev["prospetto"],
                        "--scadenze", "2026-01-31,2026-04-30,2026-07-31,2026-10-31", "--approvato", "Michele")
        self.assertEqual(code, 0, out)

    def componi(self, *args):
        code, out = run(SCRIPT, "--dir", self.dir, "--oggi", OGGI, *args)
        self.assertEqual(code, 0, out)
        return out

    def messaggio(self, uid):
        return next(m for m in self.componi()["messaggi"] if m["unita"] == uid)

    def test_una_email_per_unita_all_intestatario(self):
        out = self.componi()
        self.assertEqual([m["unita"] for m in out["messaggi"]], ["U01", "U02", "U03", "U04"])
        u02 = out["messaggi"][1]
        self.assertEqual(u02["a"], "anna.bianchi@example.it")  # mai il conduttore
        self.assertEqual(u02["oggetto"], "Condominio Prova — situazione quote 2026 — interno 2")
        self.assertTrue(out["collaudo"])

    def test_numeri_uguali_a_situazione(self):
        m = self.messaggio("U01")  # 150 € versati il 15 gennaio; rate di esempio 104,07 + 3 × 104,06
        self.assertEqual((m["dovuto"], m["versato"], m["saldo"], m["scaduto"]), (416.25, 150, 266.25, 162.19))
        c = m["corpo"]
        self.assertIn("Dovuto per l'esercizio: 416,25 €", c)
        self.assertIn("Resta da versare: 266,25 €", c)
        self.assertIn("- Rata 1 · 31 gennaio 2026 · 104,07 € · pagata", c)
        self.assertIn("- Rata 2 · 30 aprile 2026 · 104,06 € · scaduta (versati 45,93 €)", c)
        self.assertIn("- Rata 4 · 31 ottobre 2026 · 104,06 € · da pagare", c)
        self.assertIn("162,19 € riguardano rate con scadenza già passata", c)
        self.assertIn("IBAN: IT60X0542811101000000123456", c)
        self.assertIn("causale: quote 2026 interno 1", c)

    def test_solo_i_dati_della_propria_unita(self):
        for m in self.componi()["messaggi"]:
            for uid, u in self.prev.items():
                if uid != m["unita"]:
                    self.assertNotIn(u["intestatario"], m["corpo"] + m["oggetto"])

    def test_credito_e_quote_saldate(self):
        self.ris("append", "Versamenti", json.dumps({"Data": "2026-02-01", "ID unità": "U04", "Importo": 1000, "Esercizio": 2026}))
        c = self.messaggio("U04")["corpo"]
        self.assertIn("Credito a Suo favore:", c)
        self.assertNotIn("IBAN", c)  # niente coordinate se non c'è nulla da versare

    def test_registro_senza_rate(self):
        self.ris("update", "Rate", "--where", "ID unità=U03", '{"Valida": "NO"}')
        m = self.messaggio("U03")
        self.assertEqual(m["dovuto"], 0)
        self.assertIn("Le quote dell'esercizio risultano interamente versate.", m["corpo"])

    def test_unita_senza_email_e_senza_situazione(self):
        self.reg("update", "Anagrafica", "--where", "ID unità=U04", '{"Email": ""}')
        out = self.componi("--unita", "U04,U09")
        self.assertEqual(out["messaggi"], [])
        self.assertEqual([m["unita"] for m in out["senza_email"]], ["U04"])
        self.assertEqual(out["senza_situazione"], ["U09"])

    def test_salva_solo_nella_cartella_riservata(self):
        out = self.componi("--unita", "U01", "--salva")
        [path] = out["salvati"]
        self.assertEqual(os.path.dirname(path), os.path.join(self.dir, "comunicazioni"))
        with open(path, encoding="utf-8") as f:
            self.assertIn("Oggetto: Condominio Prova — situazione quote 2026 — interno 1", f.read())
        self.assertFalse(os.listdir(self.path("archivio", "2026", "comunicazioni")))
        self.assertTrue(self.componi("--unita", "U01", "--salva")["salvati"][0].endswith("_U01_2.md"))

    def test_iban_mancante_segnalato(self):
        self.reg("set", "Condominio", "Conto corrente", "")
        self.assertEqual(self.componi()["avvisi"], ["'Conto corrente' vuoto nel foglio Condominio: le email non riportano l'IBAN"])


if __name__ == "__main__":
    unittest.main()
