"""pianifica.py (0.5.0): servizio e timer systemd per il riepilogo, scritti ma mai attivati."""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

from helpers import CondominioTest, SKILLS

PIANIFICA = os.path.join(SKILLS, "automazioni", "scripts", "pianifica.py")


class TestPianifica(CondominioTest):
    def setUp(self):
        super().setUp()
        # cartella con spazi, come "Il mio Drive/Condominio Le Betulle"
        drive = os.path.join(self.dir, "Il mio Drive", "Condominio Prova")
        shutil.copytree(self.dir, drive, ignore=shutil.ignore_patterns("Il mio Drive"))
        self.dir = drive
        self.reg("set", "Condominio", "Email amministratore", "admin@example.it")
        self.conf = os.path.join(tempfile.mkdtemp(prefix="systemd-test-"), "con spazio")  # anche qui percorsi con spazi
        os.makedirs(self.conf)
        self.addCleanup(shutil.rmtree, os.path.dirname(self.conf), ignore_errors=True)
        # un finto 'claude' nel PATH, così il test non dipende dall'installazione
        self.bin = os.path.join(self.conf, "bin")
        os.makedirs(self.bin)
        with open(os.path.join(self.bin, "claude"), "w") as f:
            f.write("#!/bin/sh\nexit 0\n")
        os.chmod(os.path.join(self.bin, "claude"), 0o755)

    def pianifica(self, *args, path=None):
        env = dict(os.environ, PATH=path if path is not None else f"{self.bin}:{os.environ['PATH']}")
        p = subprocess.run([sys.executable, PIANIFICA, "--dir", self.dir,
                            "--unita-dir", os.path.join(self.conf, "unita"), "--config-dir", os.path.join(self.conf, "cfg"), *args],
                           capture_output=True, text=True, env=env)
        try:
            return p.returncode, json.loads(p.stdout)
        except ValueError:
            return p.returncode, p.stdout + p.stderr

    def leggi(self, nome):
        with open(os.path.join(self.conf, "unita", nome), encoding="utf-8") as f:
            return f.read()

    def test_file_scritti_e_comandi(self):
        code, out = self.pianifica("--giorno", "ven", "--ora", "07:30")
        self.assertEqual(code, 0, out)
        self.assertEqual(out["nome"], "riepilogo-condominio-prova")
        self.assertEqual(out["comandi"], ["systemctl --user daemon-reload",
                                          "systemctl --user enable --now riepilogo-condominio-prova.timer"])
        timer = self.leggi("riepilogo-condominio-prova.timer")
        self.assertIn("OnCalendar=Fri *-*-* 07:30:00", timer)
        self.assertIn("Persistent=true", timer)
        servizio = self.leggi("riepilogo-condominio-prova.service")
        self.assertIn(f"WorkingDirectory={self.dir}", servizio)
        self.assertIn("Type=oneshot", servizio)
        self.assertIn('--model "haiku"', servizio)
        self.assertIn(f'"{os.path.join(self.bin, "claude")}" -p', servizio)
        self.assertIn(f'Environment="PATH={self.bin}:', servizio)

    def test_solo_i_permessi_previsti(self):
        self.pianifica()
        servizio = self.leggi("riepilogo-condominio-prova.service")
        esecuzione = next(r for r in servizio.splitlines() if r.startswith("ExecStart="))
        permessi = esecuzione.split("--allowedTools", 1)[1].split()
        self.assertEqual(permessi, ['"Skill(amministratore-condominio:stato)"', '"Bash(python3', '*stato.py*)"',
                                    '"mcp__claude_ai_Gmail__send_message"'])
        self.assertNotIn("skip", servizio.lower())
        self.assertNotIn("bypass", servizio.lower())
        self.assertNotRegex(servizio, r"/cache/.*/\d+\.\d+\.\d+/")  # nessun percorso versionato del plugin

    def test_istruzioni_compilate(self):
        self.pianifica()
        with open(os.path.join(self.conf, "cfg", "riepilogo-condominio-prova.md"), encoding="utf-8") as f:
            istruzioni = f.read()
        self.assertIn("Riepilogo settimanale del Condominio Prova", istruzioni)
        self.assertIn("destinatario: admin@example.it", istruzioni)
        self.assertNotIn("{", istruzioni)
        self.assertIn(f"StandardInput=file:{os.path.join(self.conf, 'cfg', 'riepilogo-condominio-prova.md')}",
                      self.leggi("riepilogo-condominio-prova.service"))

    def test_non_sovrascrive_senza_permesso(self):
        self.pianifica()
        code, out = self.pianifica()  # identico: va bene
        self.assertEqual((code, out["aggiornati"]), (0, []))
        code, out = self.pianifica("--ora", "09:00")
        self.assertEqual(code, 2)
        self.assertIn("--sostituisci", out["errori"][0])
        self.assertIn("08:00", self.leggi("riepilogo-condominio-prova.timer"))
        code, out = self.pianifica("--ora", "09:00", "--sostituisci")
        self.assertEqual(code, 0, out)
        self.assertIn("09:00", self.leggi("riepilogo-condominio-prova.timer"))

    def test_path_senza_cartelle_dei_plugin(self):
        plugin = "/home/x/.claude/plugins/cache/condominio-agentico/amministratore-condominio/0.4.0/bin"
        self.pianifica(path=f"{self.bin}:{plugin}:/usr/bin:/usr/bin")
        servizio = self.leggi("riepilogo-condominio-prova.service")
        self.assertIn(f'Environment="PATH={self.bin}:/usr/bin"', servizio)

    def test_copia_di_sviluppo_solo_se_richiesta(self):
        self.pianifica()
        self.assertNotIn("--plugin-dir", self.leggi("riepilogo-condominio-prova.service"))
        self.pianifica("--plugin-dir", "/opt/plugin di prova", "--sostituisci")
        self.assertIn('--plugin-dir "/opt/plugin di prova" --allowedTools', self.leggi("riepilogo-condominio-prova.service"))

    def test_rimuovi_stampa_soltanto(self):
        self.pianifica()
        code, out = self.pianifica("--rimuovi")
        self.assertEqual(code, 0)
        self.assertEqual(out["comandi"][0], "systemctl --user disable --now riepilogo-condominio-prova.timer")
        self.assertIn("'", out["comandi"][1])  # percorsi con spazi tra apici
        self.assertTrue(os.path.exists(os.path.join(self.conf, "unita", "riepilogo-condominio-prova.timer")))

    def test_errori(self):
        self.reg("set", "Condominio", "Email amministratore", "")
        code, out = self.pianifica(path="/usr/bin:/bin")
        self.assertEqual(code, 2)
        self.assertEqual(len(out["errori"]), 2)
        code, out = self.pianifica("--ora", "25:00")
        self.assertNotEqual(code, 0)
        self.assertIn("Ora non valida", out)

    def test_non_chiama_mai_systemctl(self):
        spia = os.path.join(self.bin, "systemctl")
        with open(spia, "w") as f:
            f.write(f"#!/bin/sh\ntouch {self.conf}/systemctl-chiamato\n")
        os.chmod(spia, 0o755)
        self.pianifica()
        self.pianifica("--rimuovi")
        self.assertFalse(os.path.exists(os.path.join(self.conf, "systemctl-chiamato")))

    @unittest.skipUnless(shutil.which("systemd-analyze"), "systemd-analyze non disponibile")
    def test_file_validi_per_systemd(self):
        self.pianifica()
        unita = os.path.join(self.conf, "unita")
        p = subprocess.run(["systemd-analyze", "--user", "verify",
                            os.path.join(unita, "riepilogo-condominio-prova.service"),
                            os.path.join(unita, "riepilogo-condominio-prova.timer")],
                           capture_output=True, text=True)
        self.assertEqual(p.returncode, 0, p.stderr)


if __name__ == "__main__":
    unittest.main()
