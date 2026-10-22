import json,subprocess,sys,tempfile,unittest
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
SCRIPT=Path(__file__).resolve().parents[1]/'skills/software-copyright-materials/scripts/build_manual.py'
class OrderedStepsTests(unittest.TestCase):
 def test_literal_step_ordinals_restart_at_each_markdown_list(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td);(p/'a.md').write_text('# First section\n\n1. First step\n2. Second step\n\n# Next section\n\n1. Restart step\n',encoding='utf-8')
   (p/'meta.json').write_text(json.dumps({'name':'Example','version':'V1.0'}));(p/'shots.json').write_text('[]')
   res=subprocess.run([sys.executable,str(SCRIPT),'--input',str(p/'a.md'),'--metadata',str(p/'meta.json'),'--screenshots',str(p/'shots.json'),'--output',str(p/'a.docx')],capture_output=True,text=True,encoding='utf-8')
   self.assertEqual(res.returncode,0,res.stderr)
   with ZipFile(p/'a.docx') as z:doc=ET.fromstring(z.read('word/document.xml'))
   ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
   steps=[''.join(t.text or '' for t in para.findall('.//w:t',ns)) for para in doc.findall('.//w:body/w:p',ns)]
   self.assertTrue(all(s in steps for s in ['1. First step','2. Second step','1. Restart step']),steps)
if __name__=='__main__':unittest.main()
