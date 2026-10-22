import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPT=Path(__file__).resolve().parents[1]/"skills/software-copyright-materials/scripts/collect_source.py"
class SourceReviewTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base=Path(self.temp.name)
        self.root=self.base/"source"
        self.root.mkdir()
        (self.root/"src").mkdir()
        self.manifest=self.base/"manifest.json"
        self.out=self.base/"review"
    def run_tool(self,files):
        self.manifest.write_text(json.dumps({"files":files},ensure_ascii=False),encoding="utf-8")
        return subprocess.run([sys.executable,str(SCRIPT),"--root",str(self.root),"--manifest",str(self.manifest),"--out",str(self.out)],capture_output=True,text=True,encoding="utf-8")
    def test_exact_bytes_order_unicode_and_real_counts(self):
        a=("保留原文 " + "长行"*100 + "\r\n\r\n末行").encode()
        b=b"second\n"
        (self.root/"src/main.py").write_bytes(a)
        (self.root/"src/other.py").write_bytes(b)
        res=self.run_tool(["src/other.py","src/main.py"])
        self.assertEqual(res.returncode,0,res.stderr)
        inv=json.loads((self.out/"inventory.json").read_text(encoding="utf-8"))
        self.assertIs(inv["review_only"],True)
        self.assertEqual(inv["selected_file_count"],2)
        self.assertEqual(inv["selected_physical_lines"],4)
        for row,data in zip(inv["selected_files"],[b,a]):
            self.assertEqual((self.out/row["copy"]).read_bytes(),data)
            self.assertEqual(row["sha256"],hashlib.sha256(data).hexdigest())
            self.assertEqual(row["bytes"],len(data))
        self.assertEqual([r["path"] for r in inv["selected_files"]],["src/other.py","src/main.py"])
        self.assertIn(a.decode(),(self.out/"merged-review.txt").read_bytes().decode())
        self.assertNotIn(str(self.root),json.dumps(inv,ensure_ascii=False))
        self.assertEqual((self.root/"src/main.py").read_bytes(),a)
    def test_missing_input_leaves_no_review_output(self):
        self.assertNotEqual(self.run_tool(["missing.py"]).returncode,0)
        self.assertFalse(self.out.exists())
    def test_traversal_rejected(self):
        (self.base/"outside.py").write_text("outside")
        self.assertNotEqual(self.run_tool(["../outside.py"]).returncode,0)
        self.assertFalse(self.out.exists())
    def test_absolute_paths_rejected(self):
        source=self.root/"src/main.py";source.write_text("ok")
        self.assertNotEqual(self.run_tool([str(source)]).returncode,0)
        self.assertFalse(self.out.exists())
    def test_duplicates_rejected(self):
        (self.root/"src/main.py").write_text("ok")
        self.assertNotEqual(self.run_tool(["src/main.py","src/main.py"]).returncode,0)
        self.assertFalse(self.out.exists())
    def test_sensitive_config_is_not_collected(self):
        for name in [".env",".env.example","secret.key","credentials.json"]:
            (self.root/name).write_text("EXAMPLE_NOT_A_SECRET")
            self.assertNotEqual(self.run_tool([name]).returncode,0,name)
            self.assertFalse(self.out.exists())
    def test_invalid_utf8_not_silently_replaced(self):
        (self.root/"src/main.py").write_bytes(b"\xff\xfe\0")
        self.assertNotEqual(self.run_tool(["src/main.py"]).returncode,0)
        self.assertFalse(self.out.exists())
    def test_directories_and_globs_not_expanded(self):
        (self.root/"src/main.py").write_text("ok")
        for item in ["src","src/*.py"]:
            self.assertNotEqual(self.run_tool([item]).returncode,0)
            self.assertFalse(self.out.exists())
    def test_existing_output_is_never_overwritten(self):
        (self.root/"src/main.py").write_text("ok")
        self.out.mkdir();(self.out/"sentinel").write_text("keep")
        self.assertNotEqual(self.run_tool(["src/main.py"]).returncode,0)
        self.assertEqual((self.out/"sentinel").read_text(),"keep")
    def test_empty_manifest_rejected(self):
        self.assertNotEqual(self.run_tool([]).returncode,0)
        self.assertFalse(self.out.exists())
    def test_null_bytes_in_text_rejected(self):
        (self.root/"src/main.py").write_bytes(b"safe\0text")
        self.assertNotEqual(self.run_tool(["src/main.py"]).returncode,0)
        self.assertFalse(self.out.exists())
if __name__=="__main__":unittest.main()
