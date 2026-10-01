"""PPTX <-> portable, loss-aware source bundle. Requires lxml.

Preserves unknown OOXML parts instead of rebuilding them with python-pptx.
No edits: exact source bytes. Text/XML/media edits: untouched part bytes retained.
Never executes embedded content, fetches linked files, or overwrites input/output.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import shutil
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

from lxml import etree

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
MAX_BYTES = 1024 * 1024 * 1024
DRAWING_NS = {NS['a'], 'http://purl.oclc.org/ooxml/drawingml/main'}


def text_nodes(root):
    return [node for node in root.iter() if isinstance(node.tag,str)
            and etree.QName(node).localname=='t' and etree.QName(node).namespace in DRAWING_NS]


def digest(data):
    return hashlib.sha256(data).hexdigest()


def safe_part(name):
    path = PurePosixPath(name)
    if not name or '\\' in name or ':' in name or path.is_absolute() or '..' in path.parts or path.as_posix()!=name.rstrip('/'):
        raise ValueError('Unsafe package path: ' + name)
    return path


def local_file(root, name):
    path = root.joinpath(*safe_part(name).parts)
    if not path.resolve().is_relative_to(root.resolve()):
        raise ValueError('Path escapes bundle: ' + name)
    return path


def xml(data):
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    tree = etree.fromstring(data, parser)
    if tree.getroottree().docinfo.doctype:
        raise ValueError('DTD is not accepted in OOXML')
    return tree


def ordered_slides(parts):
    pres = xml(parts['ppt/presentation.xml'])
    relationships = xml(parts['ppt/_rels/presentation.xml.rels'])
    targets = {r.get('Id'): r.get('Target') for r in relationships
               if r.get('TargetMode') != 'External'}
    slides = []
    presentation_ns=etree.QName(pres).namespace
    for node in pres.findall('{'+presentation_ns+'}sldIdLst/{'+presentation_ns+'}sldId'):
        rid=next(value for key,value in node.attrib.items() if key.startswith('{') and etree.QName(key).localname=='id')
        target = targets[rid]
        name = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('ppt', target))
        safe_part(name)
        if name not in parts:
            raise ValueError('Missing slide part: ' + name)
        slides.append(name)
    return slides


def unpack(source, output):
    source, output = Path(source), Path(output)
    if output.exists():
        raise FileExistsError('Choose a new bundle directory: ' + str(output))
    with ZipFile(source) as archive:
        infos = archive.infolist()
        names = [i.filename for i in infos]
        if len(names) != len(set(names)):
            raise ValueError('Duplicate ZIP entries are not supported')
        if sum(i.file_size for i in infos) > MAX_BYTES:
            raise ValueError('Package exceeds 1 GiB uncompressed limit')
        for info in infos:
            safe_part(info.filename)
            if info.flag_bits & 1:
                raise ValueError('Encrypted ZIP is not supported')
        parts = {i.filename: archive.read(i) for i in infos if not i.is_dir()}
    slides = ordered_slides(parts)
    manifest = {
        'schema': 'pptx-source-bundle/1',
        'original': {'file': 'original.pptx', 'sha256': digest(source.read_bytes())},
        'unit': 'emu', 'slides': [], 'parts': [], 'texts': [],
        'limitations': ['Encrypted/binary .ppt are unsupported',
                        'a:t text edits only; paragraphs and run formatting are retained',
                        'External links are retained but their files are not downloaded',
                        'Editing a signed part invalidates its digital signature',
                        'Text length changes can overflow; render and review after edits',
                        'JSON summaries are inspection only; original package is authoritative'],
    }
    for name, data in parts.items():
        manifest['parts'].append({'name': name, 'sha256': digest(data), 'size': len(data)})
        if name.endswith('.xml'):
            root = xml(data)
            for index, node in enumerate(text_nodes(root)):
                manifest['texts'].append({'part': name, 'index': index,
                                           'original_text': node.text or '', 'text': node.text or ''})
        if name in slides:
            root = xml(data)
            shapes = []
            presentation_ns=etree.QName(root).namespace
            tree=root.find('.//{'+presentation_ns+'}spTree')
            for node in list(tree) if tree is not None else []:
                if etree.QName(node).localname not in {'sp','pic','graphicFrame','cxnSp','grpSp'}:continue
                nv = node.find('.//{'+presentation_ns+'}cNvPr')
                shapes.append({'type': etree.QName(node).localname,
                               'id': nv.get('id') if nv is not None else None,
                               'name': nv.get('name') if nv is not None else None,
                               'text': '\n'.join(n.text or '' for n in text_nodes(node)),
                               'xml': etree.tostring(node, encoding='unicode')})
            manifest['slides'].append({'number': slides.index(name)+1, 'part': name, 'shapes': shapes})
    manifest['slides'].sort(key=lambda s: s['number'])
    output.mkdir(parents=True)
    shutil.copyfile(source, output/'original.pptx')
    for name, data in parts.items():
        target = local_file(output/'package', name)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
    (output/'deck.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    lines = ['# PPTXのスライド内容', '', '元PPTX：'+source.name,
             'SHA-256：`'+manifest['original']['sha256']+'`', '',
             'PPTXからの抽出版．スライドの目的・実験結果の意味は推測せず，本文とnotesを記録する．',
             '文章編集はdeck.jsonのtextsで行う．このMarkdownの編集を自動でPPTXへ適用する機能はない．', '']
    for slide in manifest['slides']:
        lines += ['## Slide '+str(slide['number']), '', '部品：`'+slide['part']+'`', '', '### 本文', '']
        for shape in slide['shapes']:
            if shape['text']:
                lines += ['図形：'+str(shape['name']), '', shape['text'], '']
        rel_name = posixpath.join(posixpath.dirname(slide['part']), '_rels', posixpath.basename(slide['part'])+'.rels')
        lines += ['### 使用素材・notes', '']
        if rel_name in parts:
            for rel in xml(parts[rel_name]):
                target = rel.get('Target', '')
                if rel.get('TargetMode') == 'External':
                    lines += ['- 外部参照：'+target+'（未取得）']
                    continue
                name = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(slide['part']),target))
                if '/media/' in name or '/charts/' in name or '/diagrams/' in name:
                    lines += ['- 素材：`package/'+name+'`']
                elif '/notesSlides/' in name and name in parts:
                    values=[n.text or '' for n in text_nodes(xml(parts[name]))]
                    lines += ['', 'notes：', '', '\n'.join(values), '']
        lines += ['']
    (output/'slides.md').write_text('\n'.join(lines), encoding='utf-8')
    return manifest


def pack(bundle, output):
    bundle, output = Path(bundle), Path(output)
    if output.exists():
        raise FileExistsError('Refusing to overwrite output: ' + str(output))
    manifest = json.loads((bundle/'deck.json').read_text(encoding='utf-8'))
    if manifest.get('schema') != 'pptx-source-bundle/1':
        raise ValueError('Unsupported bundle schema')
    original = local_file(bundle, manifest['original']['file'])
    if digest(original.read_bytes()) != manifest['original']['sha256']:
        raise ValueError('Original PPTX was modified')
    edits = {}
    with ZipFile(original) as archive:
        originals = {i.filename: archive.read(i) for i in archive.infolist() if not i.is_dir()}
        listed = {p['name']: p for p in manifest['parts']}
        if len(listed) != len(manifest['parts']) or set(listed) != set(originals):
            raise ValueError('Package part list was modified')
        for name, old in originals.items():
            if digest(old) != listed[name]['sha256']:
                raise ValueError('Part baseline mismatch: ' + name)
            current = local_file(bundle/'package', name).read_bytes()
            if current != old:
                if name.endswith(('.xml', '.rels')):
                    xml(current)
                edits[name] = current
        seen, original_texts, text_roots = set(), {}, {}
        for text in manifest['texts']:
            name, index = text['part'], text['index']
            if name not in originals or type(index) is not int or index < 0:
                raise ValueError('Invalid text target')
            if (name, index) in seen:
                raise ValueError('Duplicate text target')
            seen.add((name, index))
            if name not in original_texts:
                original_texts[name] = [n.text or '' for n in text_nodes(xml(originals[name]))]
            old_texts = original_texts[name]
            if index >= len(old_texts) or old_texts[index] != text['original_text']:
                raise ValueError('Text baseline mismatch: ' + name)
            if not isinstance(text['text'], str):
                raise ValueError('Text must be a string')
            if text['text'] != text['original_text']:
                if name not in text_roots:
                    text_roots[name] = xml(edits.get(name, originals[name]))
                root = text_roots[name]
                nodes = text_nodes(root)
                if index >= len(nodes) or (nodes[index].text or '') != text['original_text']:
                    raise ValueError('Conflicting raw XML/text edits: ' + name)
                nodes[index].text = text['text']
        for name, root in text_roots.items():
            edits[name] = etree.tostring(root, encoding='UTF-8', xml_declaration=True)
        if edits and any(n.startswith('_xmlsignatures/') for n in originals):
            raise ValueError('Signed package: unchanged restore only; explicitly remove/re-sign signatures outside this tool')
        output.parent.mkdir(parents=True, exist_ok=True)
        if not edits:
            shutil.copyfile(original, output)
        else:
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, edits.get(info.filename, archive.read(info)))
    return {'changed_parts': sorted(edits), 'byte_exact': not bool(edits), 'output': str(output)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    for name in ['unpack', 'pack']:
        command = sub.add_parser(name)
        command.add_argument('input', type=Path)
        command.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    result = unpack(args.input, args.out) if args.command == 'unpack' else pack(args.input, args.out)
    print(json.dumps({'bundle': str(args.out), 'slides': len(result['slides'])} if args.command == 'unpack' else result, ensure_ascii=False))


if __name__ == '__main__':
    main()
