import os
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

SCRIPT = Path(__file__).resolve().parents[1] / "skills/software-copyright-materials/scripts/build_source_docx.py"


class BuildSourceDocxTests(unittest.TestCase):
    def test_build_source_docx_header_and_formatting(self):
        with tempfile.TemporaryDirectory() as td:
            p = Path(td)
            txt_in = p / "source.txt"
            docx_out = p / "source.docx"

            # Create 100 sample lines
            lines = [f"const a{i} = {i};\n" for i in range(100)]
            txt_in.write_text("".join(lines), encoding="utf-8")

            import subprocess
            res = subprocess.run(
                [sys.executable, str(SCRIPT), "--input", str(txt_in), "--output", str(docx_out), "--name", "测试软件", "--version", "V1.0"],
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
            self.assertEqual(res.returncode, 0, res.stderr)
            self.assertTrue(docx_out.exists())

            with ZipFile(docx_out) as z:
                # Check header
                head = "".join(z.read(x).decode() for x in z.namelist() if x.startswith("word/header") and x.endswith(".xml"))
                self.assertIn("测试软件 V1.0 源代码", head)
                self.assertIn("PAGE", head)
                self.assertIn("NUMPAGES", head)
                self.assertIn('w:val="right"', head)
                self.assertIn("<w:tbl", head)

                # Check document content
                doc = z.read("word/document.xml").decode()
                self.assertIn("const a0 = 0;", doc)
                self.assertIn("const a99 = 99;", doc)


if __name__ == "__main__":
    unittest.main()
