import json
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'src'))
from pptx_bundle import unpack, pack


def fixture(path):
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    with ZipFile(path, 'w') as z:
        z.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{p}" xmlns:r="{r}"><p:sldIdLst><p:sldId id="257" r:id="r2"/><p:sldId id="256" r:id="r1"/></p:sldIdLst></p:presentation>')
        z.writestr('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Target="slides/slide1.xml"/><Relationship Id="r2" Target="slides/slide2.xml"/></Relationships>')
        for n in [1, 2]:
            z.writestr(f'ppt/slides/slide{n}.xml', f'<p:sld xmlns:p="{p}" xmlns:a="{a}"><p:cSld><p:spTree><p:grpSp><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/></p:nvSpPr><p:txBody><a:p><a:r><a:rPr b="1"/><a:t>日本語{n}</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp></p:spTree></p:cSld><p:timing/></p:sld>')
        z.writestr('ppt/media/unknown.bin', b'unknown-media\x00\xff')
        z.writestr('ppt/embeddings/data.bin', b'embedded-workbook')


class BundleTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.source = self.root/'source.pptx'
        fixture(self.source)
        self.bundle = self.root/'bundle'
        self.manifest = unpack(self.source, self.bundle)

    def save_manifest(self):
        (self.bundle/'deck.json').write_text(json.dumps(self.manifest, ensure_ascii=False), encoding='utf-8')

    def test_no_edit_exact_and_presentation_order(self):
        self.assertEqual(self.manifest['slides'][0]['part'], 'ppt/slides/slide2.xml')
        markdown=(self.bundle/'slides.md').read_text(encoding='utf-8')
        self.assertIn('## Slide 1',markdown)
        self.assertLess(markdown.index('日本語2'),markdown.index('日本語1'))
        out = self.root/'restored.pptx'
        self.assertTrue(pack(self.bundle, out)['byte_exact'])
        self.assertEqual(out.read_bytes(), self.source.read_bytes())

    def test_text_edit_retains_other_parts_and_styles(self):
        self.manifest['texts'][0]['text'] = '変更 & <確認>'
        self.save_manifest()
        out = self.root/'edited.pptx'
        self.assertEqual(pack(self.bundle, out)['changed_parts'], ['ppt/slides/slide1.xml'])
        with ZipFile(out) as new, ZipFile(self.source) as old:
            for name in old.namelist():
                if name != 'ppt/slides/slide1.xml':self.assertEqual(new.read(name), old.read(name))
            self.assertIn(b'b="1"', new.read('ppt/slides/slide1.xml'))
            self.assertIn(b'p:timing', new.read('ppt/slides/slide1.xml'))

    def test_media_edit(self):
        (self.bundle/'package/ppt/media/unknown.bin').write_bytes(b'new-media')
        out = self.root/'media.pptx'
        self.assertEqual(pack(self.bundle, out)['changed_parts'], ['ppt/media/unknown.bin'])

    def test_no_overwrite(self):
        with self.assertRaises(FileExistsError):unpack(self.source, self.bundle)
        with self.assertRaises(FileExistsError):pack(self.bundle, self.source)

    def test_manifest_tamper(self):
        self.manifest['texts'][0]['original_text'] = 'false-baseline'
        self.save_manifest()
        with self.assertRaises(ValueError):pack(self.bundle, self.root/'bad.pptx')

    def test_path_traversal(self):
        source = self.root/'malicious.pptx'
        with ZipFile(source, 'w') as z:z.writestr('../escape', b'bad')
        with self.assertRaises(ValueError):unpack(source, self.root/'unsafe')
        self.assertFalse((self.root/'escape').exists())

    def test_strict_ooxml_namespaces(self):
        source=self.root/'strict.pptx'
        with ZipFile(self.source) as old,ZipFile(source,'w') as new:
            for name in old.namelist():
                data=old.read(name)
                if name.endswith('.xml'):
                    data=data.replace(b'http://schemas.openxmlformats.org/presentationml/2006/main',b'http://purl.oclc.org/ooxml/presentationml/main').replace(b'http://schemas.openxmlformats.org/drawingml/2006/main',b'http://purl.oclc.org/ooxml/drawingml/main').replace(b'http://schemas.openxmlformats.org/officeDocument/2006/relationships',b'http://purl.oclc.org/ooxml/officeDocument/relationships')
                new.writestr(name,data)
        manifest=unpack(source,self.root/'strict-bundle')
        self.assertEqual(len(manifest['slides']),2)
        self.assertEqual(len(manifest['texts']),2)
        pack(self.root/'strict-bundle',self.root/'strict-restored.pptx')
        self.assertEqual(source.read_bytes(),(self.root/'strict-restored.pptx').read_bytes())


if __name__ == '__main__':unittest.main()
