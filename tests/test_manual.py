import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
SCRIPT=Path(__file__).resolve().parents[1]/"skills/software-copyright-materials/scripts/build_manual.py"
class ManualTemplateTests(unittest.TestCase):
    def test_editable_word_has_real_image_slots_fields_and_unchanged_source(self):
        with tempfile.TemporaryDirectory() as td:
            p=Path(td);md=p/"manual.md";meta=p/"meta.json";shots=p/"shots.json";out=p/"manual.docx"
            body="# 示例软件说明书\n\n## 操作步骤\n\n保留 Unicode 文本与 **粗体说明**。\n\n| 操作 | 结果 |\n| --- | --- |\n| 点击 | 显示结果 |\n"
            md.write_text(body,encoding="utf-8");before=hashlib.sha256(md.read_bytes()).hexdigest()
            meta.write_text(json.dumps({"name":"示例软件","version":"V1.0 候选","owner":None,"completion_date":None,"publication_date":None},ensure_ascii=False),encoding="utf-8")
            shots.write_text(json.dumps([{"id":"图1","caption":"首页选择界面","aspect":"portrait"},{"id":"图2","caption":"结果界面","aspect":"landscape"}],ensure_ascii=False),encoding="utf-8")
            result=subprocess.run([sys.executable,str(SCRIPT),"--input",str(md),"--metadata",str(meta),"--screenshots",str(shots),"--output",str(out)],capture_output=True,text=True,encoding="utf-8")
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertEqual(hashlib.sha256(md.read_bytes()).hexdigest(),before)
            with ZipFile(out) as z:
                media=[x for x in z.namelist() if x.startswith("word/media/")]
                self.assertEqual(len(media),2)
                doc=z.read("word/document.xml").decode()
                self.assertIn("首页选择界面",doc);self.assertIn("结果界面",doc)
                self.assertIn("待确认",doc);self.assertNotIn("**粗体说明**",doc)
                head="".join(z.read(x).decode() for x in z.namelist() if x.startswith("word/header") and x.endswith(".xml"))
                self.assertIn("PAGE",head);self.assertIn("NUMPAGES",head);self.assertIn("w:val=\"right\"",head);self.assertIn("<w:tbl",head)
                self.assertNotIn(str(p),doc)
                root=ET.fromstring(doc);ns={"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                size=root.find(".//w:pgSz",ns)
                self.assertIn(int(size.attrib["{"+ns["w"]+"}w"]),range(11900,11920))
            again=subprocess.run([sys.executable,str(SCRIPT),"--input",str(md),"--metadata",str(meta),"--screenshots",str(shots),"--output",str(out)],capture_output=True)
            self.assertNotEqual(again.returncode,0)
if __name__=="__main__":unittest.main()
