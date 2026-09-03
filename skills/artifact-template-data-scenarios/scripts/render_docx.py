#!/usr/bin/env python3
"""Render Office documents to page PNGs with local font discovery and font evidence."""
import argparse
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from xml.sax.saxutils import escape


def render(source, output, soffice=None, dpi=140):
    from pdf2image import convert_from_path
    from pypdf import PdfReader
    office = soffice or shutil.which('soffice') or shutil.which('libreoffice')
    if not office:
        raise RuntimeError('LibreOffice unavailable; render/visual QA not completed')
    output.mkdir(parents=True, exist_ok=True)
    if list(output.glob('page-*.png')):
        raise ValueError('Use a fresh output directory for each render to avoid stale page images')
    directories=[Path('/System/Library/Fonts'),Path('/System/Library/Fonts/Supplemental'),Path('/Library/Fonts'),Path.home()/'Library/Fonts',Path('/Applications/wpsoffice.app/Contents/Resources/office6/fonts'),Path('/usr/share/fonts'),Path.home()/'.local/share/fonts']
    dirs=[str(p) for p in directories if p.is_dir()]
    aliases={'仿宋_GB2312':['FangS-SC','FangSong','STFangsong'],'楷体_GB2312':['HYKaiTiJ','KaiTi','Kaiti SC'],'黑体':['SimHei','Heiti SC'],'宋体':['SimSun','Songti SC']}
    with tempfile.TemporaryDirectory(prefix='scene-render-') as tmp:
        temp=Path(tmp)
        env=os.environ.copy()
        if os.name != 'nt':
            conf=['<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd"><fontconfig>']
            conf += ['<dir>'+escape(p)+'</dir>' for p in dirs]
            conf += ['<cachedir>'+escape(str(temp/'font-cache'))+'</cachedir>']
            for name, fallbacks in aliases.items():
                conf += ['<alias><family>'+name+'</family><prefer>'+''.join('<family>'+x+'</family>' for x in fallbacks)+'</prefer></alias>']
            conf += ['</fontconfig>']
            (temp/'fonts.conf').write_text(''.join(conf),encoding='utf-8')
            env['FONTCONFIG_FILE']=str(temp/'fonts.conf')
        profile=temp/'lo-profile'
        cmd=[office,'-env:UserInstallation='+profile.as_uri(),'--headless','--convert-to','pdf','--outdir',str(temp),str(source)]
        proc=subprocess.run(cmd,capture_output=True,text=True,env=env,timeout=180)
        pdf=temp/(source.stem+'.pdf')
        if proc.returncode or not pdf.exists():
            raise RuntimeError('Office rendering failed: '+proc.stdout+' '+proc.stderr)
        final_pdf=output/(source.stem+'.pdf')
        shutil.copy2(pdf,final_pdf)
    reader=PdfReader(final_pdf)
    fonts=set(); text_chars=0; cjk_chars=0
    for page in reader.pages:
        extracted=page.extract_text() or ''
        text_chars+=len(extracted)
        cjk_chars+=sum('\u4e00' <= c <= '\u9fff' for c in extracted)
        resources=page.get('/Resources',{}).get_object()
        fontdict=resources.get('/Font',{}).get_object() if hasattr(resources.get('/Font',{}),'get_object') else {}
        for f in fontdict.values():
            fonts.add(str(f.get_object().get('/BaseFont','')))
    images=convert_from_path(str(final_pdf),dpi=dpi)
    for i,page in enumerate(images,1):page.save(output/f'page-{i}.png')
    report=dict(input=str(source),pdf=str(final_pdf),pages=len(images),text_characters=text_chars,cjk_characters=cjk_chars,fonts=sorted(fonts),font_directories=dirs,render_only_fallback_candidates=aliases,visual_qa='pending_manual_review',note='DOCX font declarations were not changed. Compare actual PDF fonts to the template; fallback fonts can change pagination.')
    (output/'render-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return report


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('input',type=Path)
    p.add_argument('--output-dir',required=True,type=Path)
    p.add_argument('--soffice')
    p.add_argument('--dpi',type=int,default=140)
    a=p.parse_args()
    try:print(json.dumps(render(a.input.resolve(),a.output_dir.resolve(),a.soffice,a.dpi),ensure_ascii=False,indent=2))
    except Exception as e:p.exit(2,str(e)+'\n')


if __name__=='__main__':main()
