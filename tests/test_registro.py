"""Comportamento di registro.py (versione 0.3): il refactoring della 0.4.0 non deve cambiarlo."""
import unittest

from helpers import CondominioTest, ESERCIZIO


class TestVerifica(CondominioTest):
    def test_esempio_blocca_finche_manca_email_collaudo(self):
        code, out = self.reg("verifica")
        self.assertEqual(code, 1)
        self.assertFalse(out["ok"])
        self.assertTrue(any("Email collaudo" in p for p in out["problemi"]), out)

    def test_esempio_ok_dopo_email_collaudo(self):
        self.reg("set", "Condominio", "Email collaudo", "admin@example.it")
        code, out = self.reg("verifica")
        self.assertEqual(code, 0, out)
        self.assertTrue(out["ok"])

    def test_millesimi_che_non_quadrano(self):
        self.reg("set", "Condominio", "Email collaudo", "admin@example.it")
        self.reg("update", "Millesimi", "--where", "ID unità=U01", '{"Tab A": 290}')
        code, out = self.reg("verifica")
        self.assertEqual(code, 1)
        self.assertIn("Tab A somma a 990, non a 1000", out["problemi"])

    def test_riga_totale_riscritta_come_numeri(self):
        self.reg("update", "Millesimi", "--where", "ID unità=U01", '{"Tab A": 290}')
        ws = self.wb()["Millesimi"]
        totale = [r for r in ws.iter_rows(values_only=True) if r[0] == "TOTALE"][0]
        self.assertEqual(totale[2:5], (990, 1000, 1000))

    def test_anagrafica_e_millesimi_disallineati(self):
        self.reg("append", "Anagrafica", '{"ID unità": "U05", "Interno": "5", "Intestatario": "Nuovo"}')
        code, out = self.reg("verifica")
        self.assertIn("Unità U05 in Anagrafica ma non in Millesimi", out["problemi"])


class TestLetturaScrittura(CondominioTest):
    def test_info_conta_le_righe(self):
        code, out = self.reg("info")
        self.assertEqual(code, 0)
        self.assertEqual(out["righe_Anagrafica"], 4)
        self.assertEqual(out["righe_Millesimi"], 4)  # la riga TOTALE non conta
        self.assertEqual(out["righe_Spese"], 3)
        self.assertEqual(out["condominio"]["Modalità collaudo"], "SI")

    def test_filtri_where(self):
        _, righe = self.reg("read", "Spese", "--where", "Tabella!=A")
        self.assertEqual([r["ID"] for r in righe], [1, 2])
        _, righe = self.reg("read", "Scadenze", "--where", "ID evento=")
        self.assertEqual(righe, [])
        _, righe = self.reg("read", "Spese", "--where", "Importo=412.5")
        self.assertEqual([r["ID"] for r in righe], [1])

    def test_append_assegna_id_progressivo(self):
        code, out = self.reg("append", "Spese", '{"Data": "2026-05-01", "Fornitore": "X", "Importo": 10, "Tabella": "A"}')
        self.assertEqual(code, 0)
        self.assertEqual(out["ID"], 4)
        _, righe = self.reg("read", "Spese", "--where", "ID=4")
        self.assertEqual(righe[0]["Data"], "2026-05-01")
        self.assertEqual(righe[0]["Importo"], 10)

    def test_append_millesimi_sopra_totale(self):
        self.reg("append", "Millesimi", '{"ID unità": "U05", "Tab A": 0}')
        ws = self.wb()["Millesimi"]
        col_a = [r[0] for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
        self.assertEqual(col_a[-2:], ["U05", "TOTALE"])

    def test_colonna_sconosciuta_rifiutata(self):
        code, out = self.reg("append", "Spese", '{"Colonna inventata": 1}')
        self.assertNotEqual(code, 0)
        self.assertIn("Colonne sconosciute", out)

    def test_update_e_set(self):
        _, out = self.reg("update", "Spese", "--where", "ID=1", '{"Pagata": "SI", "Data pagamento": "2026-04-01"}')
        self.assertEqual(out["righe_aggiornate"], 1)
        _, out = self.reg("set", "Condominio", "Profilo", "professionista")
        self.assertEqual(out["Profilo"], "professionista")
        _, info = self.reg("info")
        self.assertEqual(info["condominio"]["Profilo"], "professionista")
        self.assertTrue(info["condominio"]["Ultima elaborazione"])

    def test_diario_solo_in_coda(self):
        _, prima = self.reg("read", "Diario")
        self.reg("diario", "Prova", "--dettaglio", "dettaglio", "--approvato", "Michele")
        _, dopo = self.reg("read", "Diario")
        self.assertEqual(dopo[:-1], prima)
        self.assertEqual(dopo[-1]["Operazione"], "Prova")
        self.assertEqual(dopo[-1]["Approvato da"], "Michele")


class TestSituazione(CondominioTest):
    def test_versato_e_saldo_ricalcolati(self):
        self.assertEqual(self.situazione("U03")["Versato"], 0)
        self.ris("update", "Situazione", "--where", "ID unità=U03", '{"Dovuto": 300}')
        self.ris("append", "Versamenti", '{"Data": "2026-02-01", "ID unità": "U03", "Importo": 100, "Esercizio": 2026}')
        s = self.situazione("U03")
        self.assertEqual((s["Dovuto"], s["Versato"], s["Saldo"]), (300, 100, 200))

    def test_versamento_di_altro_esercizio_non_conta(self):
        self.ris("append", "Versamenti", '{"Data": "2025-12-20", "ID unità": "U03", "Importo": 50}')
        self.ris("append", "Versamenti", f'{{"Data": "2025-12-21", "ID unità": "U03", "Importo": 70, "Esercizio": {ESERCIZIO}}}')
        self.assertEqual(self.situazione("U03")["Versato"], 70)


class TestSchema(CondominioTest):
    def test_registro_aggiornato_non_cambia(self):
        code, out = self.reg("schema")
        self.assertEqual((code, out["aggiunte"]), (0, []))
        code, out = self.ris("schema")
        self.assertEqual((code, out["aggiunte"]), (0, []))

    def test_colonna_mancante_aggiunta_in_coda_senza_perdere_dati(self):
        wb = self.wb()
        ws = wb["Spese"]
        ws.delete_cols(7)  # toglie "Quota conduttore %", come in un registro 0.2
        wb.save(self.path("registro-condominio.xlsx"))
        _, prima = self.reg("read", "Spese")
        code, out = self.reg("schema")
        self.assertEqual(out["aggiunte"], ["Spese: colonna 'Quota conduttore %' (in coda)"])
        _, dopo = self.reg("read", "Spese")
        for p, d in zip(prima, dopo):
            self.assertEqual({k: v for k, v in d.items() if k != "Quota conduttore %"}, p)


class TestRegistroVuoto(CondominioTest):
    esempio = False

    def test_anagrafica_vuota_e_solo_un_avviso(self):
        self.reg("set", "Condominio", "Email collaudo", "admin@example.it")
        code, out = self.reg("verifica")
        self.assertIn("Anagrafica vuota: inserire le unità", out["avvisi"])
        self.assertTrue(any("tutta a zero" in a for a in out["avvisi"]))


if __name__ == "__main__":
    unittest.main()
