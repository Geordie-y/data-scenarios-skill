#!/usr/bin/env python3
"""Validate sourced scene content and fill a retained Word template."""
import argparse
import ast
import json
import re
from copy import deepcopy, copy
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
from build_template import W, NS, HEADINGS, text, replace_text

KINDS = {'fact', 'proposal', 'missing', 'assumption'}
NO_CASE = '输入材料未提供相关案例'
EDITORIAL = re.compile(r'建议以|建议将|建议按|建议先|建议面向|建议由|方案建议|拟由|拟用|拟运营|材料介绍|材料列示|材料第\s*[0-9一二三四五六七八九十]+\s*页|(?:原始|输入)材料|待确认|待补充|尚未提供')


def require(value, message):
    if not value:
        raise ValueError(message)


def number(value):
    require(not isinstance(value, bool), 'Boolean is not a numeric parameter')
    result = Decimal(str(value))
    require(result.is_finite(), 'Numeric parameter must be finite')
    return result


def expression(expr, parameters):
    """Only arithmetic and declared variables; never eval user strings."""
    tree = ast.parse(expr, mode='eval')
    require(len(list(ast.walk(tree))) <= 100, 'Formula is too complex')
    names = set()
    def check(node):
        if isinstance(node, ast.Expression):
            return check(node.body)
        if isinstance(node, ast.Name):
            require(node.id in parameters, f'Unknown formula variable: {node.id}')
            names.add(node.id)
            return
        if isinstance(node, ast.Constant):
            require(isinstance(node.value, (int, float)) and not isinstance(node.value, bool), 'Only numeric constants allowed')
            number(node.value)
            return
        if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
            check(node.left); check(node.right); return
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            check(node.operand); return
        raise ValueError('Formula allows only declared variables, numbers and + - * / parentheses')
    check(tree)
    if any(parameters[n]['value'] is None for n in names):
        return None
    def calc(n):
        if isinstance(n, ast.Expression): return calc(n.body)
        if isinstance(n, ast.Name): return number(parameters[n.id]['value'])
        if isinstance(n, ast.Constant): return number(n.value)
        if isinstance(n, ast.UnaryOp): return -calc(n.operand) if isinstance(n.op, ast.USub) else calc(n.operand)
        left, right = calc(n.left), calc(n.right)
        if isinstance(n.op, ast.Add): return left + right
        if isinstance(n.op, ast.Sub): return left - right
        if isinstance(n.op, ast.Mult): return left * right
        require(right != 0, 'Division by zero')
        return left / right
    return calc(tree)


def validate(data):
    require(isinstance(data, dict), 'Root must be an object')
    require(data.get('document_mode','formal') in {'formal','review'}, 'document_mode must be formal or review')
    formal=data.get('document_mode','formal')=='formal'
    sources = data.get('sources', [])
    require(isinstance(sources, list), 'sources must be a list')
    ids = {s['id'] for s in sources}
    require(len(ids) == len(sources), 'Source IDs must be unique')
    for s in sources:
        require(all(isinstance(s.get(k), str) and s[k].strip() for k in ['id','file','locator','excerpt']), 'Each source needs id, file, locator, excerpt')
    def block(b):
        require(isinstance(b, dict), 'Content must be a block object')
        require(b.get('kind') in KINDS, 'Block kind must be fact/proposal/missing/assumption')
        require(isinstance(b.get('text'), str) and b['text'].strip(), 'Block text must not be empty')
        require('\n' not in b['text'], 'Use multiple blocks for multiple paragraphs')
        require(isinstance(b.get('sources'), list), 'Every block needs a sources list')
        require(set(b['sources']) <= ids, f'Unknown source in {b["text"][:30]}')
        if formal:
            require(b['kind']!='missing', 'Complete formal sections as designed content; put unresolved facts in limits')
            require(not EDITORIAL.search(b['text']), 'Rewrite editorial or missing-data commentary as formal business content')
        if b['kind'] == 'fact': require(b['sources'], 'Facts require source IDs')
        if b['kind'] == 'assumption': require(data.get('allow_assumptions') is True, 'Assumption requires explicit user permission recorded in allow_assumptions')
    def blocks(xs):
        require(isinstance(xs, list) and len(xs)>0, 'Required section must contain at least one block')
        for b in xs: block(b)
    require(isinstance(data.get('industries'), list) and data['industries'], 'industries must be nonempty')
    require(isinstance(data.get('limits', []), list), 'limits must be a list')
    calculations = []
    seen = set()
    for industry in data['industries']:
        require(isinstance(industry.get('title'), str) and industry['title'].strip(), 'Industry title required')
        require(industry.get('scenes'), 'Each industry needs scenes')
        for scene in industry['scenes']:
            name = scene.get('name')
            require(isinstance(name,str) and name.strip(), 'Scene name required')
            require((industry['title'],name) not in seen, 'Duplicate scene name')
            seen.add((industry['title'],name))
            for k in ['description','provider','payer']: block(scene[k])
            for k in ['background','solution','data','beneficiaries']: blocks(scene[k])
            require(isinstance(scene.get('cases'),list), 'cases must be a list; use [] when absent')
            if formal: require(scene['cases'], 'Formal material needs a sourced actual case or a clearly titled application example')
            for case in scene['cases']:
                require(case.get('title'), 'Case title required')
                block(case['detail'])
                case_type=case.get('type','actual')
                require(case_type in {'actual','example'}, 'Case type must be actual or example')
                if case_type=='actual': require(case['detail']['kind']=='fact', 'Actual cases must be sourced facts')
                else: require(case['detail']['kind'] in {'proposal','assumption'}, 'Application examples must remain marked as design in the source ledger')
            r = scene['revenue']
            for k in ['billing','subject','costs']: block(r[k])
            for k in ['period','formula','result_name','unit']:
                require(isinstance(r.get(k),str) and r[k].strip(), f'revenue.{k} required')
            params = r['parameters']
            require(isinstance(params,list) and params, 'Revenue parameters required, use value:null if missing')
            by_id = {}
            for param in params:
                require(re.fullmatch(r'[A-Za-z][A-Za-z0-9_]*', param['id']), 'Invalid parameter ID')
                require(param['id'] not in by_id, 'Duplicate parameter ID')
                require(param.get('label') and param.get('unit'), 'Parameter label/unit required')
                require(param.get('kind') in {'fact','missing','assumption'}, 'Invalid parameter kind')
                require(isinstance(param.get('sources'),list) and set(param['sources']) <= ids, 'Invalid parameter sources')
                if param['kind']=='missing': require(param.get('value') is None, 'Missing parameter must be null')
                else:
                    require(param.get('value') is not None, 'Non-missing parameter needs value')
                    number(param['value'])
                if param['kind']=='fact': require(param['sources'], 'Numeric facts require evidence')
                if param['kind']=='assumption': require(data.get('allow_assumptions') is True, 'Assumed numbers are not authorized')
                if formal: require(param['value'] is not None, 'Formal financial section needs complete calculation parameters')
                by_id[param['id']] = param
            result = expression(r['formula'], by_id)
            calculations.append(dict(scene=name, expression=r['formula'], parameters=params, period=r['period'], result=None if result is None else str(result), result_name=r['result_name'], unit=r['unit'], status='missing_parameters' if result is None else 'calculated', assumed=any(p['kind']=='assumption' for p in params)))
    return calculations


def display(b):
    return b['text']


def case_title(case):
    title=case['title']
    if case.get('type','actual')=='example' and not title.startswith('应用示例'):
        return '应用示例：'+title
    return title


def clone(p, value):
    copy = deepcopy(p)
    for el in copy.iter():
        for attr in list(el.attrib):
            if etree.QName(attr).localname in ['paraId','textId']:
                del el.attrib[attr]
    return replace_text(copy, value)


def build(data, output, template):
    results = validate(data)
    with ZipFile(template) as z:
        root = etree.fromstring(z.read('word/document.xml'))
        body = root.find('w:body',NS)
        ps = body.findall('w:p',NS)
        h1, h2 = ps[:2]
        headings = {text(p):p for p in ps if text(p) in HEADINGS}
        prototype = next(p for p in ps if text(p)=='{{section_1}}')
        missing_case = next(p for p in ps if text(p)=='{{section_6}}')
        table_prototype = body.find('w:tbl',NS)
        section = deepcopy(body.find('w:sectPr',NS))
        for el in list(body): body.remove(el)
        calc_index = 0
        for industry in data['industries']:
            body.append(clone(h1,industry['title']))
            for scene in industry['scenes']:
                body.append(clone(h2,scene['name']))
                body.append(clone(headings[HEADINGS[0]], HEADINGS[0]))
                table = deepcopy(table_prototype)
                case_summary = '；'.join(case_title(c) for c in scene['cases']) or NO_CASE
                vals=[scene['name'],display(scene['description']),display(scene['provider']),display(scene['payer']),case_summary]
                for cell,value in zip(table.findall('w:tr',NS)[1].findall('w:tc',NS),vals):
                    replace_text(cell.find('w:p',NS),value)
                body.append(table)
                section_blocks = [scene['background'],scene['solution'],[scene['provider']]+scene['data'],[scene['payer']]+scene['beneficiaries']]
                for title,blocks in zip(HEADINGS[1:5],section_blocks):
                    body.append(clone(headings[title],title))
                    for b in blocks: body.append(clone(prototype,display(b)))
                body.append(clone(headings[HEADINGS[5]],HEADINGS[5]))
                r=scene['revenue']; result=results[calc_index]; calc_index+=1
                paragraphs=[display(r['billing'])]
                param_text=[]
                for p in r['parameters']:
                    value='待补充' if p['value'] is None else str(p['value'])
                    param_text.append(f'{p["label"]} {p["id"]}＝{value} {p["unit"]}')
                paragraphs.append('测算口径：以'+display(r['subject'])+'为主体，期间为'+r['period']+'。'+'；'.join(param_text)+'。')
                equation=r['formula'].replace('*','×').replace('/','÷')
                paragraphs.append(r['result_name']+'＝'+equation+'；单位：'+r['unit']+'。')
                if result['result'] is None:
                    paragraphs[-1]+='关键参数尚未提供，暂不填写具体金额。'
                else:
                    substitutions=re.sub(r'\b[A-Za-z][A-Za-z0-9_]*\b',lambda m: str(next(p['value'] for p in r['parameters'] if p['id']==m.group())),r['formula'])
                    substitutions=substitutions.replace('*','×').replace('/','÷')
                    paragraphs[-1]=r['result_name']+'测算＝'+substitutions+'＝'+result['result']+' '+r['unit']+'。'
                paragraphs.append(display(r['costs']))
                for p in paragraphs: body.append(clone(prototype,p))
                body.append(clone(headings[HEADINGS[6]],HEADINGS[6]))
                if scene['cases']:
                    for c in scene['cases']:body.append(clone(prototype,case_title(c)+'。'+display(c['detail'])))
                else:body.append(clone(missing_case,NO_CASE))
        body.append(section)
        xml=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
        require(b'{{' not in xml, 'Unfilled template slots')
        output.parent.mkdir(parents=True,exist_ok=True)
        require(not output.exists(), 'Output exists; choose a new filename to preserve prior results')
        with ZipFile(output,'w',ZIP_DEFLATED) as dest:
            for item in z.infolist():dest.writestr(copy(item),xml if item.filename=='word/document.xml' else z.read(item.filename))
    ledger=['# 来源与编制说明','','正文按正式应用方案编写；方案设计和测算参数为编制内容，来源定位与事实核验事项保留于本清单。应用示例不代表已签约或已落地项目。','']
    for s in data['sources']:
        ledger += [f'- {s["id"]}｜{s["file"]}｜{s["locator"]}', '  摘录：'+s['excerpt']]
    ledger+=['','## 内容依据','']
    missing=[]
    def visit(obj,loc):
        if isinstance(obj,dict):
            if 'text' in obj and 'kind' in obj:
                ledger.append(f'- {loc}｜{obj["kind"]}｜依据：{", ".join(obj["sources"]) or "无原文事实引用"}｜{obj["text"]}')
                if obj['kind']=='missing': missing.append(obj['text'])
            else:
                for k,v in obj.items():visit(v,loc+'/'+str(k))
        elif isinstance(obj,list):
            for i,v in enumerate(obj,1):visit(v,loc+'/'+str(i))
    visit(data['industries'],'场景')
    ledger+=['','## 测算校验','',json.dumps(results,ensure_ascii=False,indent=2),'','## 编制依据与核验事项','']
    for line in list(dict.fromkeys(missing+data.get('limits',[]))):ledger.append('- '+line)
    ledger_path=output.with_name(output.stem+'-来源与编制说明.md')
    ledger_path.write_text('\n'.join(ledger)+'\n',encoding='utf-8')
    return dict(docx=str(output),ledger=str(ledger_path),scenes=len(results),calculations=results)


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('content',type=Path)
    p.add_argument('--out',type=Path)
    p.add_argument('--template',type=Path,default=Path(__file__).resolve().parents[1]/'assets/template.docx')
    p.add_argument('--check-only',action='store_true')
    a=p.parse_args()
    try:
        data=json.loads(a.content.read_text(encoding='utf-8'))
        if a.check_only: result=dict(valid=True,calculations=validate(data))
        else:
            require(a.out is not None,'--out is required')
            result=build(data,a.out.resolve(),a.template)
        print(json.dumps(result,ensure_ascii=False,indent=2))
    except (ValueError, KeyError, TypeError, InvalidOperation, SyntaxError) as exc:
        p.exit(2,f'Validation failed: {exc}\n')


if __name__=='__main__':main()
