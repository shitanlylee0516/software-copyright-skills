import json,subprocess,sys,tempfile,unittest
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
BASE=Path(__file__).resolve().parents[1]/'skills/software-copyright-materials/scripts'
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
class ManualFidelityTests(unittest.TestCase):
 def setUp(self):
  self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.p=Path(self.tmp.name)
  (self.p/'a.md').write_text('# 示例标题\n\n- 外层步骤\n  - 内层步骤\n- 第二外层步骤\n\n| 操作 | 结果 |\n| --- | --- |\n| **点击** | `显示结果` |\n',encoding='utf-8')
  (self.p/'meta.json').write_text(json.dumps({'name':'示例软件','version':'V1.0','owner':None}),encoding='utf-8')
  (self.p/'shots.json').write_text(json.dumps([{'id':'图 1','caption':'实际首页','aspect':'portrait'}]),encoding='utf-8')
  res=subprocess.run([sys.executable,str(BASE/'build_manual.py'),'--input',str(self.p/'a.md'),'--metadata',str(self.p/'meta.json'),'--screenshots',str(self.p/'shots.json'),'--output',str(self.p/'a.docx')],capture_output=True,text=True,encoding='utf-8')
  self.assertEqual(res.returncode,0,res.stderr)
  with ZipFile(self.p/'a.docx') as z:self.doc=ET.fromstring(z.read('word/document.xml'))
 def test_word_title_style_is_preserved(self):
  self.assertTrue(any(e.get('{'+NS['w']+'}val')=='Title' for e in self.doc.findall('.//w:pStyle',NS)))
 def test_table_contains_formatted_text_not_markdown_markers(self):
  text=''.join(x.text or '' for x in self.doc.findall('.//w:tbl//w:t',NS))
  self.assertIn('点击',text);self.assertNotIn('**',text);self.assertNotIn('`',text)
 def test_nested_steps_and_following_steps_each_survive_once(self):
  text=''.join(x.text or '' for x in self.doc.findall('.//w:t',NS))
  self.assertEqual(text.count('外层步骤'),2);self.assertEqual(text.count('内层步骤'),1);self.assertEqual(text.count('第二外层步骤'),1)
 def test_placeholder_has_helpful_alt_description(self):
  props=self.doc.findall('.//wp:docPr',NS)
  self.assertEqual(len(props),1)
  self.assertIn('实际首页',props[0].get('descr',''))
class SourceBoundaryTests(unittest.TestCase):
 def test_absolute_input_not_echoed_in_errors(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td);(p/'a.py').write_text('fictional');(p/'manifest.json').write_text(json.dumps({'files':[str(p/'a.py')]}))
   res=subprocess.run([sys.executable,str(BASE/'collect_source.py'),'--root',str(p),'--manifest',str(p/'manifest.json'),'--out',str(p/'out')],capture_output=True,text=True)
   self.assertNotEqual(res.returncode,0);self.assertNotIn(str(p),res.stderr)
 def test_sensitive_parent_directory_rejected(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td);(p/'.env').mkdir();(p/'.env'/'config.py').write_text('fictional');(p/'manifest.json').write_text(json.dumps({'files':['.env/config.py']}))
   res=subprocess.run([sys.executable,str(BASE/'collect_source.py'),'--root',str(p),'--manifest',str(p/'manifest.json'),'--out',str(p/'out')],capture_output=True)
   self.assertNotEqual(res.returncode,0);self.assertFalse((p/'out').exists())
if __name__=='__main__':unittest.main()
