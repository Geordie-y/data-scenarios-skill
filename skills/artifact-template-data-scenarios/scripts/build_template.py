#!/usr/bin/env python3
"""Distill an unchanged reference package into reusable content slots."""
import argparse
import hashlib
import json
from copy import deepcopy, copy
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
HEADINGS = ['场景一览表', '背景和痛点', '场景方案', '数据来源（数据提供方）', '目标客户（场景买单方）', '收益测算', '相关案例']


def text(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS))


def replace_text(p, value):
    first = p.find('w:r', NS)
    rpr = deepcopy(first.find('w:rPr', NS)) if first is not None and first.find('w:rPr', NS) is not None else None
    for c in list(p):
        if c.tag != f'{{{W}}}pPr':
            p.remove(c)
    run = etree.SubElement(p, f'{{{W}}}r')
    if rpr is not None:
        run.append(rpr)
    t = etree.SubElement(run, f'{{{W}}}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = value
    return p


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('reference', type=Path)
    p.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1])
    a = p.parse_args()
    assets, refs = a.root / 'assets', a.root / 'references'
    assets.mkdir(parents=True, exist_ok=True)
    refs.mkdir(parents=True, exist_ok=True)
    original = a.reference.read_bytes()
    retained = assets / 'reference.docx'
    if retained.exists() and retained.read_bytes() != original:
        raise ValueError('Retained reference differs; use an explicit new template')
    retained.write_bytes(original)
    with ZipFile(a.reference) as z:
        root = etree.fromstring(z.read('word/document.xml'))
        body = root.find('w:body', NS)
        paragraphs = body.findall('w:p', NS)
        found = [text(x) for x in paragraphs if text(x) in HEADINGS]
        if found != HEADINGS:
            raise ValueError('Reference does not contain the expected seven-section contract')
        replace_text(paragraphs[0], '{{industry}}')
        replace_text(paragraphs[1], '{{scene_name}}')
        current = None
        kept = set()
        for el in list(body)[2:]:
            if el.tag == f'{{{W}}}p':
                t = text(el)
                if t in HEADINGS:
                    current = HEADINGS.index(t)
                elif current is not None and current not in kept:
                    replace_text(el, '{{section_' + str(current) + '}}')
                    kept.add(current)
                else:
                    body.remove(el)
        table = body.find('w:tbl', NS)
        for i, cell in enumerate(table.findall('w:tr', NS)[1].findall('w:tc', NS)):
            ps = cell.findall('w:p', NS)
            replace_text(ps[0], '{{overview_' + str(i) + '}}')
            for extra in ps[1:]:
                cell.remove(extra)
        output = assets / 'template.docx'
        with ZipFile(output, 'w', ZIP_DEFLATED) as dest:
            for item in z.infolist():
                dest.writestr(copy(item), etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True) if item.filename == 'word/document.xml' else z.read(item.filename))
        def node(e):
            if e is None:
                return None
            return dict(tag=etree.QName(e).localname, attrs={etree.QName(k).localname:v for k,v in e.attrib.items()}, children=[node(c) for c in e])
        styles = etree.fromstring(z.read('word/styles.xml'))
        evidence = dict(reference_sha256=hashlib.sha256(original).hexdigest(), section=node(body.find('w:sectPr', NS)), styles=[node(s) for s in styles.findall('w:style', NS) if s.get(f'{{{W}}}styleId') in ['1','2','3','4','10']], table=node(table.find('w:tblPr', NS)), grid=node(table.find('w:tblGrid', NS)), numbering=node(etree.fromstring(z.read('word/numbering.xml'))), parts={n:dict(bytes=len(z.read(n)), sha256=hashlib.sha256(z.read(n)).hexdigest(), editable=n=='word/document.xml') for n in z.namelist() if not n.endswith('/')})
        (refs/'style-evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2),encoding='utf-8')
    print(output)


if __name__ == '__main__':
    main()
