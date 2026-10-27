import streamlit as st
import fitz, re, zipfile, io, os, glob, copy
import xml.etree.ElementTree as ET
from lxml import etree as LET
from collections import Counter

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('',NS)
def q(tag): return '{'+NS+'}'+tag

def sheet_data(blob):
    z=zipfile.ZipFile(io.BytesIO(blob))
    ss=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        root=ET.fromstring(z.read('xl/sharedStrings.xml'))
        ss=[''.join(x.text or '' for x in el.iter(q('t'))) for el in root.findall(q('si'))]
    root=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    data={}
    for c in root.iter(q('c')):
        v=c.find(q('v')); inline=c.find(q('is'))
        if c.get('t')=='s' and v is not None: value=ss[int(v.text)]
        elif c.get('t')=='inlineStr' and inline is not None:value=''.join(x.text or '' for x in inline.iter(q('t')))
        elif v is not None:value=v.text or ''
        else:value=''
        data[c.get('r')]=value
    return data

def write_cells(blob, changes):
    """Patch cells while preserving Office XML namespaces and package structure."""
    namespace_fixes = {
        'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
        'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
        'xr2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
        'xr3': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',
    }
    zin = zipfile.ZipFile(io.BytesIO(blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w') as zout:
        for item in zin.infolist():
            name = item.filename
            if name == 'xl/calcChain.xml':
                continue  # Excel rebuilds the calculation chain
            raw = zin.read(name)
            if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                root = LET.fromstring(raw)
                ignorable = root.get('{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable', '')
                missing = {prefix: namespace_fixes[prefix] for prefix in ignorable.split()
                           if prefix not in root.nsmap and prefix in namespace_fixes}
                if missing:
                    nsmap = dict(root.nsmap)
                    nsmap.update(missing)
                    newroot = LET.Element(root.tag, nsmap=nsmap)
                    newroot.attrib.update(root.attrib)
                    newroot.text = root.text
                    for child in list(root):
                        root.remove(child)
                        newroot.append(child)
                    root = newroot
                if name == 'xl/worksheets/sheet1.xml':
                    byaddr = {c.get('r'): c for c in root.iter(q('c'))}
                    for addr, value in changes.items():
                        c = byaddr.get(addr)
                        if c is None:
                            continue
                        if c.find(q('f')) is not None:
                            continue  # never overwrite a formula
                        for child in list(c):
                            if child.tag in (q('v'), q('is')):
                                c.remove(child)
                        c.set('t', 'inlineStr')
                        node = LET.SubElement(c, q('is'))
                        text_node = LET.SubElement(node, q('t'))
                        text_node.text = str(value)
                raw = LET.tostring(root, encoding='UTF-8', xml_declaration=True)
            elif name == 'xl/_rels/workbook.xml.rels':
                root = LET.fromstring(raw)
                for rel in list(root):
                    if rel.get('Type', '').endswith('/calcChain'):
                        root.remove(rel)
                raw = LET.tostring(root, encoding='UTF-8', xml_declaration=True)
            elif name == '[Content_Types].xml':
                root = LET.fromstring(raw)
                for child in list(root):
                    if child.get('PartName') == '/xl/calcChain.xml':
                        root.remove(child)
                raw = LET.tostring(root, encoding='UTF-8', xml_declaration=True)
            zout.writestr(item, raw)
    result = out.getvalue()
    with zipfile.ZipFile(io.BytesIO(result)) as z:
        if z.testzip() is not None:
            raise ValueError('ZIP integrity check failed')
        for name in z.namelist():
            if name.endswith(('.xml', '.rels')):
                LET.fromstring(z.read(name))
    return result

def normalize(x):return re.sub(r'[^A-Z0-9]','',str(x).upper())
def pdf_details(data):
    doc=fitz.open(stream=data,filetype='pdf')
    pages=[p.get_text(sort=True) for p in doc]
    head='\n'.join(pages[:2]); m=re.search(r'STYLE\s*:\s*(\d{4,6})',head,re.I)
    style=m.group(1) if m else ''
    sections=[];colors=[]
    for n,txt in enumerate(pages):
        if 'MARKETING BOM' in txt.upper() or 'TRIMS' in txt.upper() or 'THREADS' in txt.upper():sections.append(n+1)
        for line in txt.splitlines():
            if re.search(r'(ECO SUNDRY|VSW|PGD WASH)',line,re.I):
                s=re.split(r'\s+-\s+(?:ECO SUNDRY|VSW|PGD WASH)',line,flags=re.I)[0].strip()
                if 2<len(s)<45 and not re.search(r'(STYLE|CONTENT|COMPONENT|MATERIAL|THREAD|PAGE|STATUS|ON-PRODUCT)',s,re.I):colors.append(s)
    return style,pages,sections,list(dict.fromkeys(colors))

def learn(records):
    mapping={}; yy={}; item_info={}
    for name,blob in records:
        cells=sheet_data(blob)
        for r in range(20,150):
            desc=cells.get(f'B{r}','').strip(); color=cells.get(f'C{r}','').strip(); trim=cells.get(f'D{r}','').strip()
            if desc:
                code=re.match(r'\s*(\d{1,6})\s*[- ]',desc)
                if code:item_info.setdefault(code.group(1),[]).append((name,r,desc,cells.get(f'F{r}',''),cells.get(f'G{r}','')))
            if color and trim and not color.upper().startswith(('ALL CC','BODY COLOR')):
                if re.search(r'\([A-Z0-9-]{4,}\)',trim):
                    mapping.setdefault(normalize(color),Counter())[trim]+=1
            if desc and cells.get(f'F{r}',''):
                code=re.match(r'\s*(\d{1,6})\s*[- ]',desc)
                if code:yy.setdefault(code.group(1),Counter())[str(cells[f'F{r}'])]+=1
    return mapping,yy,item_info

st.set_page_config(page_title='AERIE TRIM LIST Generator',layout='wide')
st.title('AERIE Tech Pack → TRIM LIST')
st.caption('기존 TRIM LIST 서식을 보존하고 과거 부자재 코드 및 TRIM YY를 참고값으로 자동 제안하는 검토용 도구')
files=st.file_uploader('Tech Pack PDF 업로드',type=['pdf'])
base=st.file_uploader('출력 양식으로 사용할 기존 TRIM LIST Excel 업로드',type=['xlsx'])
refs=st.file_uploader('과거 TRIM LIST (복수 선택 가능)',type=['xlsx'],accept_multiple_files=True)
if files and base:
    style,pages,sections,colors=pdf_details(files.getvalue())
    st.write('Tech Pack Style:',style or '미인식','| BOM 관련 페이지:',', '.join(map(str,sections)))
    st.write('추출된 컬러 후보 (확인 필요):',', '.join(colors[:30]) or '없음')
    records=[(base.name,base.getvalue())]+[(f.name,f.getvalue()) for f in refs]
    mapping,yy,info=learn(records)
    original=sheet_data(base.getvalue())
    template_style=original.get('C5','')
    st.warning(f'템플릿 Style: {template_style} / Tech Pack Style: {style}. 기존 템플릿에 다른 Style의 행이 남아 있을 수 있으므로 반드시 검토하세요.') if normalize(template_style)!=normalize(style) else st.success('템플릿과 Tech Pack의 Style 번호가 일치합니다.')
    updates={'C5':style,'A1':style} if style else {}
    candidates=[]; last_desc=''; last_code=''; mismatches=[]
    for r in range(20,150):
        desc=original.get(f'B{r}','').strip()
        if desc:
            last_desc=desc; m=re.match(r'\s*(\d{1,6})\s*[- ]',desc);last_code=m.group(1) if m else ''
        body=original.get(f'C{r}','').strip();trim=original.get(f'D{r}','').strip()
        if last_code and desc and last_code in yy and not original.get(f'F{r}',''):
            val=yy[last_code].most_common(1)[0][0];updates[f'F{r}']=val;candidates.append((r,'TRIM YY',val,'과거 동일 품목 코드'))
        if body and not body.upper().startswith(('ALL CC','BODY COLOR')):
            key=normalize(body)
            if key in mapping and (not trim or not re.search(r'\([A-Z0-9-]{4,}\)',trim)):
                val=mapping[key].most_common(1)[0][0]
                updates[f'D{r}']=val;candidates.append((r,'TRIM COLOR',val,'과거 동일 바디 컬러'))
    st.subheader('자동 제안 결과')
    st.write(f'자동 입력 후보: {len(candidates)}건. 기존 데이터는 기본적으로 보존합니다.')
    if candidates:st.dataframe([{'Row':r,'Field':field,'Value':v,'Source':source} for r,field,v,source in candidates],hide_index=True)
    st.info('현재 버전은 템플릿 기존 행에 대해 코드를 추천하는 1차 버전입니다. Tech Pack의 신규 품목/컬러 행을 자동 추가하거나 BOM 전체를 완전 대조하지 않습니다.')
    if st.button('검토용 Excel 생성',type='primary'):
        result=write_cells(base.getvalue(),updates)
        st.download_button('TRIM LIST Excel 다운로드',data=result,file_name=f'{style or "NEW"}_TRIM_LIST_DRAFT.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
else:
    st.info('Tech Pack PDF와 기준 TRIM LIST Excel을 올려주세요. 과거 발주서도 함께 올리면 코드 추천이 확장됩니다.')
