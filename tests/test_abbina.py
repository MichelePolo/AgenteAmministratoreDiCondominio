"""abbina.py (0.4.0): proposta di abbinamento dei bonifici e registrazione dei versamenti confermati."""
import json
import os
import unittest

from helpers import CondominioTest, RIPARTO, SKILLS, run

ABBINA = os.path.join(SKILLS, "versamenti", "scripts", "abbina.py")
OGGI = "2026-09-26"
SCADENZE = "2026-01-31,2026-04-30,2026-07-31,2026-10-31"


class BaseAbbina(CondominioTest):
    def setUp(self):
        super().setUp()
        # rate approvate: U03 (Luca Verdi) ha 4 rate da 150,25 €; nessun versamento
        _, prev = self.riparto("--rate", 4, "--titolo", "Preventivo 2026")
        self.rate = prev["per_unita"]
        code, out = run(RIPARTO, "approva", "--dir", self.dir, "--prospetto", prev["prospetto"],
                        "--scadenze", SCADENZE, "--approvato", "Michele")
        self.assertEqual(code, 0, out)

    def proponi(self, movimenti):
        path = self.path("movimenti.json")
        with open(path, "w") as f:
            json.dump(movimenti, f)
        code, out = run(ABBINA, "--dir", self.dir, "--oggi", OGGI, "--movimenti", path)
        self.assertEqual(code, 0, out)
        return out["movimenti"]

    def registra(self, righe):
        path = self.path("confermati.json")
        with open(path, "w") as f:
            json.dump(righe, f)
        return run(ABBINA, "registra", "--dir", self.dir, "--oggi", OGGI, "--versamenti", path, "--approvato", "Michele")

    def bonifico(self, ordinante, causale, importo, data="2026-09-10"):
        return {"data": data, "importo": importo, "ordinante": ordinante, "causale": causale}


class TestAbbina(BaseAbbina):

    def test_nome_interno_e_rata_proposti(self):
        rata1 = self.rate["U03"]["rate"][0]
        [m] = self.proponi([self.bonifico("VERDI LUCA", "CONDOMINIO RATA 1 INT 3", rata1)])
        self.assertEqual((m["esito"], m["unita"], m["rata"]), ("proposto", "U03", 1))
        self.assertEqual(m["motivi"], ["nome completo (Luca Verdi)", "interno 3 nella causale", "importo uguale alla rata 1"])
        v = m["versamento"]
        self.assertEqual((v["ID unità"], v["Importo"], v["Esercizio"], v["Rata"]), ("U03", rata1, 2026, 1))

    def test_registrato_non_si_ripropone(self):
        mov = [self.bonifico("VERDI LUCA", "rata 1 int 3", 150)]
        [m] = self.proponi(mov)
        code, out = self.registra([m["versamento"]])
        self.assertEqual(code, 0, out)
        [m] = self.proponi(mov)
        self.assertEqual(m["esito"], "gia_registrato")
        # anche con la causale trascritta diversamente dallo stesso PDF
        [m] = self.proponi([self.bonifico("Verdi  Luca", "RATA 1 - INTERNO 3", "150,00")])
        self.assertEqual(m["esito"], "gia_registrato")

    def test_nessun_indizio_da_abbinare(self):
        [m] = self.proponi([self.bonifico("SCONOSCIUTO SRL", "CONDOMINIO", 77.77)])
        self.assertEqual((m["esito"], m["motivo"], m["candidati"]), ("da_abbinare", "nessun indizio", []))
        self.assertNotIn("versamento", m)

    def test_solo_cognome_non_basta(self):
        [m] = self.proponi([self.bonifico("VERDI", "quota condominio", 99)])
        self.assertEqual((m["esito"], m["motivo"]), ("da_abbinare", "indizi deboli"))
        self.assertEqual(m["candidati"][0]["unita"], "U03")

    def test_omonimi_da_abbinare(self):
        self.reg("append", "Anagrafica", '{"ID unità": "U05", "Interno": "5", "Intestatario": "Mario Rossi"}')
        [m] = self.proponi([self.bonifico("ROSSI MARIO", "quota condominio", 99)])
        self.assertEqual((m["esito"], m["motivo"]), ("da_abbinare", "più unità possibili"))
        self.assertEqual(sorted(c["unita"] for c in m["candidati"]), ["U01", "U05"])
        # l'interno nella causale scioglie l'omonimia
        [m] = self.proponi([self.bonifico("ROSSI MARIO", "quota condominio int. 5", 99)])
        self.assertEqual((m["esito"], m["unita"]), ("proposto", "U05"))

    def test_paga_il_conduttore(self):
        [m] = self.proponi([self.bonifico("GIALLI PAOLO", "spese condominiali", 50)])
        self.assertEqual((m["esito"], m["unita"]), ("proposto", "U02"))
        self.assertEqual(m["motivi"], ["nome completo (Paolo Gialli)"])

    def test_addebito_non_e_un_versamento(self):
        """Fino alla 0.5 gli addebiti si ignoravano; dalla 0.6 si abbinano alle spese (test_addebiti.py)."""
        [m] = self.proponi([self.bonifico("ENEL ENERGIA", "addebito SDD", -412.5)])
        self.assertEqual((m["esito"], m["id_spesa"]), ("spesa_pagata", 1))
        self.assertNotIn("versamento", m)

    def test_due_bonifici_uguali_coprono_rate_successive(self):
        rata = self.rate["U03"]["rate"][0]
        mov = [self.bonifico("VERDI LUCA", "int 3", rata), self.bonifico("VERDI LUCA", "int 3", rata)]
        a, b = self.proponi(mov)
        self.assertNotEqual(a["id_movimento"], b["id_movimento"])
        self.assertEqual((a["rata"], b["rata"]), (1, 2))

    def test_unita_indicata_dall_utente(self):
        [m] = self.proponi([{"data": "2026-09-20", "importo": 150, "unita": "U03"}])
        self.assertEqual((m["esito"], m["unita"], m["motivi"]), ("proposto", "U03", ["unità indicata dall'utente"]))
        self.assertEqual(m["rata"], 1)

    def test_formati_italiani_e_errori(self):
        [m] = self.proponi([{"data": "10/09/2026", "importo": "1.234,56", "unita": "U01"}])
        self.assertEqual((m["data"], m["importo"]), ("2026-09-10", 1234.56))
        path = self.path("m.json")
        with open(path, "w") as f:
            json.dump([{"data": "ieri", "importo": 10}], f)
        code, out = run(ABBINA, "--dir", self.dir, "--movimenti", path)
        self.assertEqual(code, 2)
        self.assertIn("Movimento 1: data non riconosciuta", out["errori"][0])


class TestRegistra(BaseAbbina):
    def test_registra_aggiorna_situazione_e_diario_anonimo(self):
        rata = self.rate["U03"]["rate"][0]
        [m] = self.proponi([self.bonifico("VERDI LUCA", "rata 1 int 3", rata)])
        code, out = self.registra([m["versamento"]])
        self.assertEqual(code, 0, out)
        self.assertEqual(out["registrati"], 1)
        self.assertEqual(self.situazione("U03")["Versato"], rata)
        _, diario = self.reg("read", "Diario")
        ultima = diario[-1]
        self.assertEqual(ultima["Operazione"], "Registrato 1 versamento")
        riga = json.dumps(ultima, ensure_ascii=False)
        for vietato in ("Verdi", "U03", "€", str(rata).replace(".", ",")):
            self.assertNotIn(vietato, riga)

    def test_registra_rifiuta_duplicati_e_unita_inesistenti(self):
        [m] = self.proponi([self.bonifico("VERDI LUCA", "int 3", 10)])
        code, out = self.registra([m["versamento"], m["versamento"]])
        self.assertEqual(code, 2)
        self.assertIn("già registrato", out["errori"][0])
        code, out = self.registra([{"Data": "2026-09-01", "ID unità": "U99", "Importo": 10}])
        self.assertIn("unità U99 inesistente", out["errori"][0])
        _, vers = self.ris("read", "Versamenti")
        self.assertEqual(len(vers), 2)  # solo i 2 di esempio: nessuna scrittura parziale

    def test_crea_la_riga_di_situazione_mancante(self):
        code, out = self.registra([{"Data": "2027-01-15", "ID unità": "U04", "Importo": 100, "Esercizio": 2027}])
        self.assertEqual(code, 0, out)
        self.assertEqual(out["situazioni_create"], ["U04"])
        _, righe = self.ris("read", "Situazione", "--where", "Esercizio=2027")
        self.assertEqual((righe[0]["ID unità"], righe[0]["Versato"]), ("U04", 100))


if __name__ == "__main__":
    unittest.main()
