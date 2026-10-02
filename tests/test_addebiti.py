"""Addebiti dell'estratto conto (0.6.0): spese pagate, spese da abbinare, nuove spese, registrazione."""
import json
import unittest

from helpers import run
from test_abbina import ABBINA, OGGI, BaseAbbina

# Spese di esempio: 1 ENEL 412,50 (14/3, non pagata), 2 Otis 780 (pagata a mano il 20/2, senza ID movimento),
# 3 Studio Tecnico Bruni 1250 (5/4, non pagata).


class TestAddebiti(BaseAbbina):
    def addebito(self, descrizione, causale, importo, data="2026-04-02"):
        return self.bonifico(descrizione, causale, -abs(importo), data=data)

    def registra_obj(self, **dati):
        path = self.path("confermati.json")
        with open(path, "w") as f:
            json.dump(dati, f)
        return run(ABBINA, "registra", "--dir", self.dir, "--oggi", OGGI, "--versamenti", path, "--approvato", "Michele")

    def spesa(self, sid):
        _, righe = self.reg("read", "Spese", "--where", f"ID={sid}")
        return righe[0]

    def test_addebito_paga_la_spesa(self):
        [m] = self.proponi([self.addebito("ENEL ENERGIA SPA", "ADDEBITO SDD FATTURA MARZO", 412.5)])
        self.assertEqual((m["esito"], m["id_spesa"]), ("spesa_pagata", 1))
        self.assertEqual((m["data"], m["data_spesa"]), ("2026-04-02", "2026-03-14"))  # la data del movimento resta sua
        self.assertEqual(m["motivi"], ["stesso importo", "fornitore ENEL nella descrizione", "data entro 180 giorni dal documento"])
        code, out = self.registra_obj(spese_pagate=[m["spesa_pagata"]])
        self.assertEqual((code, out["spese_pagate"]), (0, [1]), out)
        s = self.spesa(1)
        self.assertEqual((s["Pagata"], s["Data pagamento"], s["ID movimento"]), ("SI", "2026-04-02", m["id_movimento"]))
        [m] = self.proponi([self.addebito("ENEL ENERGIA SPA", "ADDEBITO SDD FATTURA MARZO", 412.5)])
        self.assertEqual(m["esito"], "gia_registrato")

    def test_spesa_pagata_a_mano_si_collega_senza_doppioni(self):
        [m] = self.proponi([self.addebito("OTIS SERVIZI", "BONIFICO MANUTENZIONE", 780, data="2026-02-20")])
        self.assertEqual((m["esito"], m["id_spesa"], m["gia_pagata"]), ("spesa_pagata", 2, True))
        self.registra_obj(spese_pagate=[m["spesa_pagata"]])
        s = self.spesa(2)
        self.assertEqual((s["Data pagamento"], s["ID movimento"]), ("2026-02-20", m["id_movimento"]))  # data a mano intatta

    def test_due_spese_con_lo_stesso_importo(self):
        for f in ("Idraulico Rossi", "Elettricista Verdi"):
            self.reg("append", "Spese", json.dumps({"Data": "2026-03-01", "Fornitore": f, "Importo": 100, "Tabella": "A", "Esercizio": 2026}))
        [m] = self.proponi([self.addebito("BONIFICO A FAVORE DI", "PAGAMENTO FATTURA", 100)])
        self.assertEqual((m["esito"], m["motivo"]), ("spesa_da_abbinare", "più spese con lo stesso importo"))
        self.assertEqual(sorted(c["id_spesa"] for c in m["candidati"]), [4, 5])
        # il fornitore nella descrizione scioglie il dubbio
        [m] = self.proponi([self.addebito("ROSSI IDRAULICO", "PAGAMENTO FATTURA", 100)])
        self.assertEqual((m["esito"], m["id_spesa"]), ("spesa_pagata", 4))

    def test_uno_a_uno_nello_stesso_elenco(self):
        a, b = self.proponi([self.addebito("ENEL ENERGIA", "SDD", 412.5), self.addebito("ENEL ENERGIA", "SDD", 412.5, data="2026-04-03")])
        self.assertEqual((a["esito"], b["esito"]), ("spesa_pagata", "nuova_spesa"))

    def test_fuori_finestra_e_senza_fornitore(self):
        [m] = self.proponi([self.addebito("PAGAMENTO", "DISPOSIZIONE", 412.5, data="2026-12-31")])
        self.assertEqual((m["esito"], m["motivo"]), ("spesa_da_abbinare", "importo uguale ma fornitore e data non tornano"))

    def test_commissioni_nuova_spesa_bancaria(self):
        [m] = self.proponi([self.addebito("COMMISSIONI TENUTA CONTO", "", 4.5, data="2026-09-20")])
        self.assertEqual((m["esito"], m["spesa_bancaria"]), ("nuova_spesa", True))
        r = m["nuova_spesa"]
        self.assertEqual((r["Tabella"], r["Quota conduttore %"], r["Pagata"], r["Importo"]), ("A", 0, "SI", 4.5))
        self.assertIn("documento non previsto", r["Note"])
        code, out = self.registra_obj(nuove_spese=[r])
        self.assertEqual((code, out["nuove_spese"]), (0, [4]), out)
        s = self.spesa(4)
        self.assertEqual((s["Pagata"], s["Data pagamento"], s["Tabella"]), ("SI", "2026-09-20", "A"))

    def test_nuova_spesa_senza_tabella_rifiutata(self):
        [m] = self.proponi([self.addebito("IDRAULICO BIANCHI SRL", "PAGAMENTO", 150)])
        r = m["nuova_spesa"]
        self.assertIsNone(r["Tabella"])
        self.assertIn("documento da archiviare", r["Note"])
        code, out = self.registra_obj(nuove_spese=[r])
        self.assertEqual(code, 2)
        self.assertIn("manca la Tabella", out["errori"][0])
        code, out = self.registra_obj(nuove_spese=[dict(r, Tabella="A", **{"Quota conduttore %": 0})])
        self.assertEqual(code, 0, out)

    def test_interessi_creditori_ignorati(self):
        [m] = self.proponi([self.bonifico("INTERESSI CREDITORI", "", 0.12)])
        self.assertEqual(m["esito"], "ignorato")
        self.assertIn("movimento della banca", m["motivo"])

    def test_registrazione_mista_tutto_o_niente(self):
        rata = self.rate["U03"]["rate"][0]
        movs = self.proponi([self.bonifico("VERDI LUCA", "rata 1 int 3", rata),
                             self.addebito("ENEL ENERGIA", "SDD", 412.5),
                             self.addebito("COMMISSIONI", "", 2)])
        v = next(m["versamento"] for m in movs if m["esito"] == "proposto")
        p = next(m["spesa_pagata"] for m in movs if m["esito"] == "spesa_pagata")
        n = next(m["nuova_spesa"] for m in movs if m["esito"] == "nuova_spesa")
        code, out = self.registra_obj(versamenti=[v], spese_pagate=[p, dict(p, **{"ID spesa": 99})], nuove_spese=[n])
        self.assertEqual(code, 2)
        self.assertIn("spesa 99 inesistente", " ".join(out["errori"]))
        self.assertEqual(self.situazione("U03")["Versato"], 0)  # niente scritto
        self.assertEqual(self.spesa(1)["Pagata"], "NO")
        code, out = self.registra_obj(versamenti=[v], spese_pagate=[p], nuove_spese=[n])
        self.assertEqual(code, 0, out)
        _, diario = self.reg("read", "Diario")
        self.assertEqual(diario[-1]["Operazione"], "Registrato 1 versamento, 1 spesa segnata pagata, 1 nuova spesa pagata")


class TestSenzaDocumento(BaseAbbina):
    def test_avviso_dopo_30_giorni_anche_nel_riepilogo(self):
        [m] = self.proponi([self.bonifico("IDRAULICO BIANCHI", "PAGAMENTO", -150, data="2026-08-01")])
        path = self.path("c.json")
        with open(path, "w") as f:
            json.dump({"nuove_spese": [dict(m["nuova_spesa"], Tabella="A")]}, f)
        run(ABBINA, "registra", "--dir", self.dir, "--oggi", OGGI, "--versamenti", path, "--approvato", "Michele")
        _, out = self.reg("--oggi", "2026-08-20", "verifica")
        self.assertFalse(any("senza documento" in x for x in out["avvisi"]))
        _, out = self.reg("--oggi", "2026-09-26", "verifica")
        self.assertIn("Spese pagate da più di 30 giorni ma senza documento in archivio: ID [4]", out["avvisi"])
        from test_stato import STATO
        code, testo = run(STATO, "--dir", self.dir, "--oggi", "2026-09-26", "--testo")
        self.assertIn("· 1 spesa pagata senza documento in archivio da oltre 30 giorni", testo)


if __name__ == "__main__":
    unittest.main()
