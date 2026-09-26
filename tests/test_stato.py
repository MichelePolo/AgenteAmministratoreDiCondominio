"""stato.py (0.4.0): cruscotto in sola lettura e azioni suggerite in ordine di urgenza."""
import datetime as dt
import json
import os
import shutil
import unittest

from helpers import CondominioTest, SKILLS, run

STATO = os.path.join(SKILLS, "stato", "scripts", "stato.py")
OGGI = dt.date(2026, 9, 26)


class BaseStato(CondominioTest):
    def setUp(self):
        super().setUp()
        self.reg("set", "Condominio", "Email collaudo", "admin@example.it")  # nessun problema bloccante di partenza

    def stato(self, oggi=OGGI):
        code, out = run(STATO, "--dir", self.dir, "--oggi", oggi.isoformat())
        self.assertEqual(code, 0, out)
        return out

    def azioni_testo(self, out):
        return [a["testo"] for a in out["azioni"]]

    def scadenza(self, giorni, tipo="assemblea", descrizione="Assemblea ordinaria", **extra):
        riga = {"Data": (OGGI + dt.timedelta(days=giorni)).isoformat(), "Tipo": tipo, "Descrizione": descrizione, **extra}
        self.reg("append", "Scadenze", json.dumps(riga))


class TestStato(BaseStato):

    def test_sola_lettura(self):
        files = [self.path("registro-condominio.xlsx"), self.path("registro-riservato.xlsx")]
        prima = [os.path.getmtime(f) for f in files]
        self.stato()
        self.assertEqual([os.path.getmtime(f) for f in files], prima)

    def test_problema_bloccante_prima_di_tutto(self):
        self.reg("set", "Condominio", "Email collaudo", "")
        for i in range(3):
            open(self.path("da analizzare", f"doc{i}.pdf"), "w").close()
        out = self.stato()
        self.assertEqual(out["azioni"][0]["priorita"], 1)
        self.assertIn("Email collaudo", out["azioni"][0]["testo"])

    def test_convocazione_da_inviare(self):
        self.scadenza(12)
        out = self.stato()
        self.assertEqual(out["azioni"][0]["priorita"], 2)
        self.assertIn("la convocazione va inviata entro il 3 ottobre 2026", out["azioni"][0]["testo"])

    def test_convocazione_fuori_termine(self):
        self.scadenza(3)
        a = self.stato()["azioni"][0]
        self.assertEqual(a["priorita"], 1)
        self.assertIn("termine di 5 giorni è già passato", a["testo"])

    def test_convocazione_gia_nel_diario(self):
        self.scadenza(12)
        self.reg("diario", "Inviata convocazione a 4 destinatari", "--approvato", "Michele")
        out = self.stato()
        self.assertTrue(out["assemblea"]["convocazione_inviata"])
        self.assertFalse(any("convocazione" in t for t in self.azioni_testo(out)))

    def test_rendiconto_entro_180_giorni(self):
        self.reg("set", "Condominio", "Fine esercizio", "2025-12-31")
        out = self.stato(dt.date(2026, 5, 15))
        self.assertEqual(out["assemblea"]["termine_rendiconto"], "2026-06-29")
        a = out["azioni"][0]
        self.assertEqual(a["priorita"], 2)
        self.assertIn("entro il 29 giugno 2026 (mancano 45 giorni)", a["testo"])

    def test_rate_scadute_e_documenti_vecchi(self):
        self.ris("append", "Rate", json.dumps({"ID unità": "U03", "Esercizio": 2026, "Rata": 1, "Scadenza": "2026-06-30", "Importo": 200}))
        doc = self.path("da analizzare", "vecchio.pdf")
        open(doc, "w").close()
        dieci_giorni_fa = (dt.datetime.combine(OGGI, dt.time()) - dt.timedelta(days=10)).timestamp()
        os.utime(doc, (dieci_giorni_fa, dieci_giorni_fa))
        out = self.stato()
        self.assertEqual(out["pagamenti"]["unita_in_ritardo"], 1)
        self.assertEqual(out["pagamenti"]["dettaglio"][0]["unita"], "U03")
        self.assertEqual([a["priorita"] for a in out["azioni"]], [3, 3])
        self.assertIn("1 unità ha rate scadute per 200,00 €", out["azioni"][0]["testo"])
        self.assertIn("il più vecchio caricato 10 giorni fa", out["azioni"][1]["testo"])

    def test_documento_recente_meno_urgente(self):
        open(self.path("da analizzare", "nuovo.pdf"), "w").close()
        [a] = self.stato(dt.date.today())["azioni"]
        self.assertEqual((a["priorita"], a["testo"]), (4, "1 documento da archiviare"))

    def test_al_massimo_tre_azioni_in_ordine(self):
        self.reg("set", "Condominio", "Email collaudo", "")
        self.scadenza(12)
        self.ris("append", "Rate", json.dumps({"ID unità": "U03", "Esercizio": 2026, "Rata": 1, "Scadenza": "2026-06-30", "Importo": 200}))
        open(self.path("da analizzare", "doc.pdf"), "w").close()
        out = self.stato()
        self.assertEqual([a["priorita"] for a in out["azioni"]], [1, 2, 3])

    def test_scadenze_ricorrenti_da_riproporre(self):
        self.scadenza(-20, tipo="manutenzione", descrizione="Verifica ascensore", Ricorrenza="biennale")
        out = self.stato()
        self.assertEqual([x["descrizione"] for x in out["scadenze"]["da_riproporre"]], ["Verifica ascensore"])
        self.scadenza(700, tipo="manutenzione", descrizione="Verifica ascensore", Ricorrenza="biennale")
        self.assertEqual(self.stato()["scadenze"]["da_riproporre"], [])

    def test_bozze_dimenticate(self):
        os.makedirs(self.path("prospetti"), exist_ok=True)
        open(self.path("prospetti", "2026-06-01_riparto_preventivo_bozza.xlsx"), "w").close()
        out = self.stato()
        self.assertEqual(out["bozze"], [{"file": "prospetti/2026-06-01_riparto_preventivo_bozza.xlsx", "giorni": 117}])
        self.assertIn("in bozza da più di due mesi", self.azioni_testo(out)[-1])

    def test_registro_riservato_mancante(self):
        shutil.move(self.path("registro-riservato.xlsx"), self.path("altrove.xlsx"))
        out = self.stato()
        self.assertIsNone(out["pagamenti"])
        self.assertTrue(any("Registro riservato non trovato" in a for a in out["registro"]["avvisi"]))

    def test_cartella_sbagliata(self):
        code, out = run(STATO, "--dir", self.path("archivio"))
        self.assertNotEqual(code, 0)
        self.assertIn("la cartella di lavoro non è quella del condominio", out)



class TestTesto(BaseStato):
    """stato.py --testo: il riepilogo della chat e dell'email programmata."""

    def testo(self, oggi=OGGI):
        code, out = run(STATO, "--dir", self.dir, "--oggi", oggi.isoformat(), "--testo")
        self.assertEqual(code, 0, out)
        return out

    def test_tutto_in_ordine(self):
        t = self.testo()
        self.assertTrue(t.startswith("Condominio Prova — riepilogo al 26 settembre 2026\n\n"))
        self.assertIn("· Registro in ordine · Modalità collaudo attiva", t)
        self.assertTrue(t.endswith("Tutto in ordine."))
        self.assertNotIn("Cosa fare adesso", t)

    def test_righe_e_azioni(self):
        self.scadenza(12, descrizione="Assemblea ordinaria")
        self.scadenza(20, tipo="manutenzione", descrizione="Verifica ascensore")
        self.ris("append", "Rate", json.dumps({"ID unità": "U03", "Esercizio": 2026, "Rata": 1, "Scadenza": "2026-06-30", "Importo": 200}))
        for n in ("a.pdf", "b.pdf"):
            open(self.path("da analizzare", n), "w").close()
        t = self.testo(dt.date.today())  # i file appena creati hanno la data di oggi
        self.assertIn("· 2 documenti da archiviare\n", t)
        self.assertIn("Prossime scadenze: Verifica ascensore il", t)
        self.assertIn("(convocazione non ancora inviata)", t)
        self.assertIn("· Rate scadute: 1 unità, 200,00 €", t)
        self.assertIn("Cosa fare adesso:\n1. ", t)
        self.assertIn('→ "archivia i documenti"', t)

    def test_nessun_dato_per_unita(self):
        self.ris("append", "Rate", json.dumps({"ID unità": "U03", "Esercizio": 2026, "Rata": 1, "Scadenza": "2026-06-30", "Importo": 200}))
        t = self.testo()
        _, anagrafica = self.reg("read", "Anagrafica")
        for u in anagrafica:
            for campo in ("ID unità", "Intestatario", "Email", "Conduttore"):
                if u.get(campo):
                    self.assertNotIn(str(u[campo]), t, campo)

    def test_problemi_e_riservato_mancante(self):
        self.reg("set", "Condominio", "Email collaudo", "")
        shutil.move(self.path("registro-riservato.xlsx"), self.path("altrove.xlsx"))
        t = self.testo()
        self.assertIn("· Pagamenti non verificati: registro riservato non trovato", t)
        self.assertIn("· Registro: 1 problema da correggere", t)

    def test_rendiconto_da_convocare(self):
        self.reg("set", "Condominio", "Fine esercizio", "2025-12-31")
        self.assertIn("· Rendiconto: assemblea da tenere entro il 29 giugno 2026", self.testo(dt.date(2026, 5, 15)))


if __name__ == "__main__":
    unittest.main()
