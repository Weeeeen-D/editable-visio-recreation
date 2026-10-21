"""Static VSDX checks. This is not a Visio renderer or full schema validator."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import posixpath
import re
import zipfile
from pathlib import Path
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET

V = '{http://schemas.microsoft.com/office/visio/2012/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'


def cells(element):
    return {c.get('N'): c for c in element.findall(V + 'Cell')} if element is not None else {}


def number(cell):
    try:
        value = float(cell.get('V'))
        return value if math.isfinite(value) else None
    except (AttributeError, TypeError, ValueError):
        return None


def rel_target(source, target):
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc:
        raise ValueError('Internal relationship contains a URL')
    path = unquote(parsed.path)
    if '\\' in path:
        raise ValueError('Backslash in OPC target')
    resolved = posixpath.normpath(path.lstrip('/') if path.startswith('/') else
                                 posixpath.join(posixpath.dirname(source), path))
    if resolved in ('.', '..') or resolved.startswith('../'):
        raise ValueError('Relationship escapes package root')
    return resolved


def inspect(path, profile=None, allow_media=False):
    path = Path(path)
    report = {'file': str(path.resolve()), 'static_ok': False, 'errors': [], 'warnings': [],
              'pages': [], 'shapes': 0, 'text_shapes': 0, 'connections': 0,
              'unverified': ['Actual Visio opening/rendering', 'Visual fidelity and text clipping',
                             'Interactive resizing and connector following',
                             'Master/style inheritance and font substitution']}
    errors, warnings = report['errors'], report['warnings']
    profile = profile or {}
    try:
        report['sha256'] = hashlib.sha256(path.read_bytes()).hexdigest()
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            names = {i.filename for i in infos}
            if len(names) != len(infos):
                errors.append('Duplicate ZIP entries')
            if sum(i.file_size for i in infos) > 256 * 1024 * 1024:
                raise ValueError('Expanded package exceeds the 256 MiB inspection limit')
            bad = z.testzip()
            if bad:
                errors.append('ZIP CRC failure: ' + bad)
            xml = {}
            for name in sorted(names):
                if name.endswith(('.xml', '.rels')):
                    data = z.read(name)
                    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
                        errors.append('DTD/entity not allowed: ' + name)
                        continue
                    try:
                        xml[name] = ET.fromstring(data)
                    except ET.ParseError as exc:
                        errors.append(f'Malformed XML: {name}: {exc}')
            for required in ('[Content_Types].xml', '_rels/.rels', 'visio/document.xml',
                             'visio/pages/pages.xml', 'visio/pages/_rels/pages.xml.rels'):
                if required not in xml:
                    errors.append('Missing or invalid required XML: ' + required)
            relationships = {}
            for name, root in xml.items():
                if not name.endswith('.rels'):
                    continue
                source = '' if name == '_rels/.rels' else posixpath.join(
                    posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
                ids = {}
                if root.tag != REL + 'Relationships':
                    errors.append('Invalid relationship namespace: ' + name)
                for rel in root:
                    rid = rel.get('Id')
                    if not rid or rid in ids:
                        errors.append('Missing/duplicate relationship ID: ' + name)
                    ids[rid] = None
                    if rel.get('TargetMode') == 'External':
                        warnings.append('External relationship not fetched: ' + name + ' ' + str(rid))
                        continue
                    try:
                        target = rel_target(source, rel.attrib['Target'])
                        ids[rid] = target
                        if target not in names:
                            errors.append('Missing relationship target: ' + target)
                    except (KeyError, ValueError) as exc:
                        errors.append(f'Invalid relationship in {name}: {exc}')
                relationships[source] = ids
            report['media_parts'] = sorted(n for n in names if n.startswith('visio/media/'))
            report['embedded_parts'] = sorted(n for n in names if n.startswith('visio/embeddings/'))
            if (report['media_parts'] or report['embedded_parts']) and not allow_media:
                errors.append('Image/embedded-object parts found; use --allow-media only for an approved exception')
            if any('vbaproject' in n.lower() for n in names):
                errors.append('Unexpected macro payload in VSDX')
            index = xml.get('visio/pages/pages.xml')
            if index is not None:
                entries = index.findall(V + 'Page')
                if not entries:
                    errors.append('No pages')
                for entry in entries:
                    rid_node = entry.find(V + 'Rel')
                    target = relationships.get('visio/pages/pages.xml', {}).get(
                        rid_node.get(R + 'id') if rid_node is not None else None)
                    if target not in xml:
                        errors.append('Page relationship cannot be resolved: ' + str(entry.get('ID')))
                        continue
                    sheet = cells(entry.find(V + 'PageSheet'))
                    width, height = number(sheet.get('PageWidth')), number(sheet.get('PageHeight'))
                    details = {'id': entry.get('ID'), 'part': target, 'width_mm': None,
                               'height_mm': None, 'shape_count': 0}
                    for key, value, limit in [('width_mm', width, 'max_width_mm'),
                                              ('height_mm', height, 'max_height_mm')]:
                        if value is None or value <= 0:
                            errors.append(f'{target}: missing/invalid page {key}')
                        else:
                            details[key] = round(value * 25.4, 6)
                            if limit in profile and value * 25.4 > profile[limit] + 0.01:
                                errors.append(f'{target}: page {key} exceeds profile')
                    page = xml[target]
                    if page.tag != V + 'PageContents':
                        errors.append('Invalid PageContents namespace: ' + target)
                    shapes = page.findall('.//' + V + 'Shape')
                    ids = [s.get('ID') for s in shapes]
                    if None in ids or len(ids) != len(set(ids)):
                        errors.append('Missing/duplicate shape IDs: ' + target)
                    details['shape_count'] = len(shapes)
                    report['shapes'] += len(shapes)
                    report['pages'].append(details)
                    for link in page.findall('.//' + V + 'Connect'):
                        report['connections'] += 1
                        if link.get('FromSheet') not in ids or link.get('ToSheet') not in ids:
                            errors.append('Connection points to a missing shape: ' + target)
                    for shape in shapes:
                        label = f'{target} shape {shape.get("ID")}'
                        if shape.get('Type') == 'Foreign' and not allow_media:
                            errors.append(label + ': foreign object is not editable primitives')
                        text = shape.find(V + 'Text')
                        if text is not None and ''.join(text.itertext()).strip():
                            report['text_shapes'] += 1
                            rows = shape.findall(f'{V}Section[@N="Character"]/{V}Row')
                            if profile and not rows:
                                warnings.append(label + ': text formatting is inherited; check in Visio')
                            content = ''.join(text.itertext())
                            for row in rows:
                                formatting = cells(row)
                                size = number(formatting.get('Size'))
                                if 'font_pt' in profile:
                                    if size is None:
                                        warnings.append(label + ': inherited/unresolved font size')
                                    elif abs(size * 72 - profile['font_pt']) > 0.01:
                                        errors.append(label + ': font size differs from profile')
                                font_cell = formatting.get('Font')
                                font = font_cell.get('V', '') if font_cell is not None else ''
                                if profile and len(rows) == 1:
                                    expected = profile.get('cjk_font' if re.search(r'[\u3400-\u9fff]', content)
                                                           else 'western_font')
                                    if not font or font.isdigit():
                                        warnings.append(label + ': unresolved font name')
                                    elif expected and font.casefold() != expected.casefold():
                                        errors.append(label + ': font name differs from profile')
                                elif profile:
                                    warnings.append(label + ': mixed text runs require font review in Visio')
                        # Master-based connectors often cache absolute geometry; do not flag it as static.
                        if not shape.get('Master') and 'BeginX' not in cells(shape):
                            fixed = [c for sec in shape.findall(f'{V}Section[@N="Geometry"]')
                                     for row in sec.findall(V + 'Row') if row.get('T') in ('MoveTo', 'LineTo')
                                     for c in row.findall(V + 'Cell') if c.get('N') in ('X', 'Y')
                                     and number(c) not in (None, 0) and c.get('F') in (None, '', 'No Formula')]
                            if fixed:
                                warnings.append(label + ': fixed geometry; verify visible resizing')
            if report['shapes'] == 0:
                errors.append('No editable page shapes found')
    except (OSError, zipfile.BadZipFile, ValueError, RuntimeError) as exc:
        errors.append(str(exc))
    report['warnings'] = list(dict.fromkeys(warnings))
    report['static_ok'] = not errors
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('file', type=Path)
    parser.add_argument('--profile', default='source')
    parser.add_argument('--profiles', type=Path, default=Path(__file__).resolve().parents[1] / 'references/profiles.json')
    parser.add_argument('--allow-media', action='store_true')
    parser.add_argument('--strict', action='store_true', help='Treat unresolved warnings as a failed static check')
    parser.add_argument('--report', type=Path)
    args = parser.parse_args()
    profiles = json.loads(args.profiles.read_text(encoding='utf-8-sig'))
    if args.profile not in profiles:
        parser.error('Unknown profile: ' + args.profile)
    result = inspect(args.file, profiles[args.profile], args.allow_media)
    output = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        if args.report.resolve() == args.file.resolve():
            parser.error('Report cannot replace the input file')
        args.report.write_text(output + '\n', encoding='utf-8')
    print(output)
    return 0 if result['static_ok'] and not (args.strict and result['warnings']) else 1


if __name__ == '__main__':
    raise SystemExit(main())
