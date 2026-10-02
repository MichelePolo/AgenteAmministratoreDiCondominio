"""Memoria di chi paga per chi (0.6.0): abbina.py ricorda / dimentica e il peso nel punteggio."""
import json
import unittest

from helpers import run
from test_abbina import ABBINA, OGGI, BaseAbbina


class TestMemoria(BaseAbbina):
    def ricorda(self, ordinante, unita):
        return run(ABBINA, "ricorda", "--dir", self.dir, "--oggi", OGGI, "--ordinante", ordinante,
                   "--unita", unita, "--approvato", "Michele")

    def dimentica(self, ordinante, *extra):
        return run(ABBINA, "dimentica", "--dir", self.dir, "--oggi", OGGI, "--ordinante", ordinante,
                   "--approvato", "Michele", *extra)

    def ferrari(self, causale="CONDOMINIO BETULLE"):
        [m] = self.proponi([self.bonifico("FERRARI GIOVANNA", causale, 120)])
        return m

    def test_senza_memoria_nessun_indizio(self):
        m = self.ferrari()
        self.assertEqual((m["esito"], m["motivo"]), ("da_abbinare", "nessun indizio"))

    def test_ordinante_ricordato_basta_da_solo(self):
        code, out = self.ricorda("Ferrari Giovanna", "U03")
        self.assertEqual(code, 0, out)
        self.assertEqual((out["ordinante"], out["unita"], out["anche_per"]), ("FERRARI GIOVANNA", "U03", []))
        m = self.ferrari()
        self.assertEqual((m["esito"], m["unita"]), ("proposto", "U03"))
        self.assertEqual(m["motivi"], ["ordinante ricordato per l'interno 3"])

    def test_ricordato_due_volte_non_duplica(self):
        self.ricorda("FERRARI GIOVANNA", "U03")
        code, out = self.ricorda("ferrari  giovanna", "U03")
        self.assertTrue(out["gia_ricordato"])
        _, righe = self.ris("read", "Ordinanti")
        self.assertEqual(len(righe), 1)

    def test_stesso_ordinante_per_due_unita(self):
        self.ricorda("FERRARI GIOVANNA", "U03")
        code, out = self.ricorda("FERRARI GIOVANNA", "U04")
        self.assertEqual(out["anche_per"], ["U03"])
        self.assertIn("resteranno da abbinare", out["avviso"])
        m = self.ferrari()
        self.assertEqual((m["esito"], m["motivo"]), ("da_abbinare", "più unità possibili"))
        self.assertEqual(sorted(c["unita"] for c in m["candidati"]), ["U03", "U04"])
        m = self.ferrari("CONDOMINIO INT 4")  # l'interno nella causale scioglie il dubbio
        self.assertEqual((m["esito"], m["unita"]), ("proposto", "U04"))

    def test_dimentica_disattiva_senza_cancellare(self):
        self.ricorda("FERRARI GIOVANNA", "U03")
        code, out = self.dimentica("FERRARI GIOVANNA")
        self.assertEqual((code, out["disattivate"]), (0, ["U03"]))
        self.assertEqual(self.ferrari()["motivo"], "nessun indizio")
        _, righe = self.ris("read", "Ordinanti")
        self.assertEqual([(r["ID unità"], r["Attiva"]) for r in righe], [("U03", "NO")])
        code, out = self.dimentica("FERRARI GIOVANNA")
        self.assertEqual(code, 2)
        # riattivare riusa la riga
        code, out = self.ricorda("FERRARI GIOVANNA", "U03")
        self.assertTrue(out["riattivato"])
        _, righe = self.ris("read", "Ordinanti")
        self.assertEqual([(r["ID unità"], r["Attiva"]) for r in righe], [("U03", "SI")])

    def test_dimentica_una_sola_unita(self):
        self.ricorda("FERRARI GIOVANNA", "U03")
        self.ricorda("FERRARI GIOVANNA", "U04")
        code, out = self.dimentica("FERRARI GIOVANNA", "--unita", "U04")
        self.assertEqual(out["disattivate"], ["U04"])
        self.assertEqual(self.ferrari()["unita"], "U03")

    def test_errori(self):
        code, out = self.ricorda("FERRARI GIOVANNA", "U99")
        self.assertEqual(code, 2)
        self.assertIn("unità U99 inesistente", out["errori"][0])
        code, out = self.ricorda(" . ", "U03")
        self.assertIn("ordinante vuoto", out["errori"][0])

    def test_diario_senza_nomi(self):
        self.ricorda("FERRARI GIOVANNA", "U03")
        self.dimentica("FERRARI GIOVANNA")
        _, diario = self.reg("read", "Diario")
        ultime = json.dumps(diario[-2:], ensure_ascii=False)
        self.assertIn("Memorizzato 1 ordinante", ultime)
        self.assertIn("Disattivato 1 ordinante", ultime)
        for vietato in ("FERRARI", "Ferrari", "U03"):
            self.assertNotIn(vietato, ultime)


if __name__ == "__main__":
    unittest.main()
