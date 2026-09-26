"""Frontmatter delle skill: nome uguale alla cartella, description entro 250 caratteri (limite di visualizzazione)."""
import glob
import os
import re
import unittest

from helpers import SKILLS


class TestSkillMd(unittest.TestCase):
    def test_frontmatter(self):
        files = sorted(glob.glob(os.path.join(SKILLS, "*", "SKILL.md")))
        self.assertGreaterEqual(len(files), 7)
        for f in files:
            with open(f, encoding="utf-8") as h:
                testo = h.read()
            cartella = os.path.basename(os.path.dirname(f))
            with self.subTest(skill=cartella):
                self.assertTrue(testo.startswith("---\n"))
                self.assertIn(f"\nname: {cartella}\n", testo)
                blocco = re.search(r"\ndescription: >\n((?:  .*\n)+)", testo)
                self.assertIsNotNone(blocco, "description mancante o non in forma '>'")
                descrizione = " ".join(r.strip() for r in blocco.group(1).splitlines())
                self.assertLessEqual(len(descrizione), 250, descrizione)


if __name__ == "__main__":
    unittest.main()
