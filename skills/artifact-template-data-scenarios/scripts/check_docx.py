#!/usr/bin/env python3
"""Verify source-derived structure, immutable package parts and overview/body consistency."""
import argparse
import json
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from build_template import W, NS, HEADINGS, text
from generate_docx import validate, display, NO_CASE, case_title, EDITORIAL


def check(path, content, template):
    data=json.loads(content.read_text(encoding='utf-8'))
    validate(data)
    failures=[]
    with ZipFile(path) as z, ZipFile(template) as baseline:
        for name in baseline.namelist():
            if name!='word/document.xml' and z.read(name)!=baseline.read(name):
                failures.append('Preserve-only package part changed: '+name)
        root=etree.fromstring(z.read('word/document.xml'))
        ref=etree.fromstring(baseline.read('word/document.xml'))
        body=root.find('w:body',NS)
        paragraphs=body.findall('w:p',NS)
        actual=[text(p) for p in paragraphs if p.find('w:pPr/w:pStyle',NS) is not None and p.find('w:pPr/w:pStyle',NS).get(f'{{{W}}}val')=='4']
        scenes=[s for i in data['industries'] for s in i['scenes']]
        if actual!=HEADINGS*len(scenes):failures.append('Seven-section headings/order mismatch')
        if '{{' in ''.join(root.xpath('//w:t/text()',namespaces=NS)):failures.append('Unfilled slots')
        if data.get('document_mode','formal')=='formal' and EDITORIAL.search(''.join(root.xpath('//w:t/text()',namespaces=NS))):failures.append('Editorial language remains in formal document')
        tables=body.findall('w:tbl',NS)
        if len(tables)!=len(scenes):failures.append('Overview table count mismatch')
        for index,(table,scene) in enumerate(zip(tables,scenes),1):
            rows=table.findall('w:tr',NS)
            expected=[scene['name'],display(scene['description']),display(scene['provider']),display(scene['payer']),'；'.join(case_title(c) for c in scene['cases']) or NO_CASE]
            if len(rows)!=2 or [text(c) for c in rows[1].findall('w:tc',NS)]!=expected:failures.append(f'Overview {index} differs from content')
            if [text(c) for c in rows[0].findall('w:tc',NS)]!=['场景名称','场景描述','数据来源','目标客户（买单方）','相关案例']:failures.append('Table header mismatch')
            ref_table=ref.find('w:body/w:tbl',NS)
            for prop in ['w:tblPr','w:tblGrid']:
                if etree.tostring(table.find(prop,NS))!=etree.tostring(ref_table.find(prop,NS)):failures.append('Table geometry changed')
        if etree.tostring(body.find('w:sectPr',NS))!=etree.tostring(ref.find('w:body/w:sectPr',NS)):failures.append('Page geometry changed')
    return dict(passed=not failures,scenes=len(scenes),failures=failures,visual_qa='separate_required',facts='source references validated; factual support requires human/model review')


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('docx',type=Path)
    p.add_argument('--content',required=True,type=Path)
    p.add_argument('--template',type=Path,default=Path(__file__).resolve().parents[1]/'assets/template.docx')
    a=p.parse_args()
    result=check(a.docx,a.content,a.template)
    print(json.dumps(result,ensure_ascii=False,indent=2))
    if not result['passed']:raise SystemExit(1)


if __name__=='__main__':main()
