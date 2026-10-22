import contextlib,importlib.util,io,json,sys,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
SCRIPT=Path(__file__).resolve().parents[1]/'skills/software-copyright-materials/scripts/collect_source.py'
class SourceErrorPrivacyTests(unittest.TestCase):
 def test_read_failure_does_not_echo_private_absolute_path(self):
  with tempfile.TemporaryDirectory() as td:
   p=Path(td);(p/'a.py').write_text('fictional');(p/'manifest.json').write_text(json.dumps({'files':['a.py']}))
   spec=importlib.util.spec_from_file_location('collector_for_test',SCRIPT);mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
   err=io.StringIO()
   with patch.object(sys,'argv',[str(SCRIPT),'--root',str(p),'--manifest',str(p/'manifest.json'),'--out',str(p/'out')]),patch.object(Path,'read_bytes',side_effect=PermissionError('private path '+str(p/'a.py'))),contextlib.redirect_stderr(err):
    with self.assertRaises(SystemExit):mod.main()
   self.assertNotIn(str(p),err.getvalue());self.assertFalse((p/'out').exists())
if __name__=='__main__':unittest.main()
