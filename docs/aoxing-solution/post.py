import docx, copy, re, sys
from docx.shared import Pt, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
import json,os
PM=json.load(open('pagemap.json')) if os.path.exists('pagemap.json') else {}
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
src,dst=sys.argv[1],sys.argv[2]
d=docx.Document(src)
body=d.element.body
HEADER='奥星石化物料平衡与日效益 · 解决方案思路'
# ---------- page setup
s=d.sections[0]
s.page_width=Twips(11906); s.page_height=Twips(16838)
s.top_margin=Twips(1440); s.bottom_margin=Twips(1300); s.left_margin=Twips(1270); s.right_margin=Twips(1270)
s.header_distance=Twips(708); s.footer_distance=Twips(708)
s.different_first_page_header_footer=True
# ---------- styles
def setfont(style,east,size=None,color=None,bold=None,ascii_='Calibri'):
    rpr=style.element.get_or_add_rPr()
    rf=rpr.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts'); rpr.insert(0,rf)
    for a in ('w:ascii','w:hAnsi','w:cs'): rf.set(qn(a),ascii_)
    rf.set(qn('w:eastAsia'),east)
    for a in ('w:asciiTheme','w:eastAsiaTheme','w:hAnsiTheme'):
        if rf.get(qn(a)) is not None: del rf.attrib[qn(a)]
    if size: style.font.size=Pt(size)
    if color: style.font.color.rgb=RGBColor.from_string(color)
    if bold is not None: style.font.bold=bold
st=d.styles
# docDefaults east asian font
dd=d.styles.element.find(qn('w:docDefaults')).find(qn('w:rPrDefault')).find(qn('w:rPr'))
rf=dd.find(qn('w:rFonts')); rf.set(qn('w:eastAsia'),'宋体')
lang=dd.find(qn('w:lang'))
if lang is None: lang=OxmlElement('w:lang'); dd.append(lang)
lang.set(qn('w:eastAsia'),'zh-CN')
setfont(st['Normal'],'宋体',10.5)
st['Normal'].paragraph_format.line_spacing=1.3
st['Normal'].paragraph_format.space_after=Pt(6)
setfont(st['Heading 1'],'微软雅黑',16,'1F3864',True)
setfont(st['Heading 2'],'微软雅黑',12.5,'1F7A8C',True)
setfont(st['Heading 3'],'微软雅黑',12,'1F4D78',True)
st['Heading 1'].paragraph_format.space_before=Pt(18); st['Heading 1'].paragraph_format.space_after=Pt(12)
st['Heading 2'].paragraph_format.space_before=Pt(14); st['Heading 2'].paragraph_format.space_after=Pt(6)
st['Heading 3'].paragraph_format.space_before=Pt(10); st['Heading 3'].paragraph_format.space_after=Pt(4)
for n in ['List Paragraph','Quote']:
    if n in [x.name for x in st]: setfont(st[n],'宋体')
# ---------- remove title block, promote headings
ps=d.paragraphs
assert ps[0].text.startswith('奥星石化') and ps[3].text.startswith('说明')
summary=ps[2].text; note=ps[3].text
for p in ps[:4]: p._element.getparent().remove(p._element)
lvl={'Heading 2':'Heading 1','Heading 3':'Heading 2','Heading 4':'Heading 3'}
heads=[]
for p in d.paragraphs:
    n=p.style.name
    if n in lvl:
        p.style=st[lvl[n]]
        if lvl[n] in ('Heading 1','Heading 2'): heads.append((lvl[n],p.text))
    # chapter headings start on new page
for p in d.paragraphs:
    if p.style.name=='Heading 1': p.paragraph_format.page_break_before=True
for p in list(d.paragraphs):
    if p.text.strip() in ('图 1 平台总体架构','图 2 分阶段演进路线与门槛'): p._element.getparent().remove(p._element)
# ---------- formulas
F={
 'm = (V_t':[('m = (V',0),('t',1),(' − V',0),('w',1),(') × VCF',0),('20',1),(' × (ρ',0),('20',1),(' − 0.0011) − m',0),('roof',1)],
 '\\text{原料}':[('原料 + 辅料及工艺水 − 产品 − 自用 − 排放 − 损失 − Δ罐存 = 平衡差',0)],
 'P_{net}':[('P',0),('net',1),(' = P',0),('tax',1),(' / (1 + r',0),('VAT',1),(') − T',0),('consumption',1),(' × k',0),('t/L',1),(' − T',0),('surcharge',1)],
 '\\Pi_{\\text{全厂}}':[('Π',0),('全厂',1),(' = ΣQ',0),('p',1),('P',0),('p',1),(' − ΣQ',0),('r',1),('P',0),('r',1),(' − C',0),('energy',1),(' − C',0),('aux',1),(' − C',0),('fix',1),(' + ΔInv',0)],
 '\\Pi_{\\text{装置}}':[('Π',0),('装置',1),(' = ΣQ',0),('out',1),('P',0),('out',1),(' − ΣQ',0),('in',1),('P',0),('in',1),(' − C',0),('energy',1),(' − C',0),('aux',1)],
 '\\text{价格影响}':[('价格影响 = Q',0),('今日',1),(' × (P',0),('今日',1),(' − P',0),('昨日',1),(');  数量影响 = P',0),('昨日',1),(' × (Q',0),('今日',1),(' − Q',0),('昨日',1),(')',0)],
}
nf=0
for p in d.paragraphs:
    if p.style.name!='Code': continue
    key=[k for k in F if p.text.startswith(k)]
    assert key, p.text
    for r in list(p.runs): r._element.getparent().remove(r._element)
    p.style=st['Normal']; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(8)
    for txt,sub in F[key[0]]:
        r=p.add_run(txt); r.font.name='Cambria Math'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'宋体'); r.font.size=Pt(11.5); r.italic=not re.search('[一-鿿]',txt)
        if sub: r.font.subscript=True
    nf+=1
print('formulas',nf)
# ---------- cover + TOC at start
first=d.paragraphs[0]
def para_before(text='',size=None,bold=False,color=None,align='center',east='微软雅黑',space_after=None,br=False):
    p=first.insert_paragraph_before()
    p.alignment={'center':WD_ALIGN_PARAGRAPH.CENTER,'left':WD_ALIGN_PARAGRAPH.LEFT}[align]
    if text:
        r=p.add_run(text); r.bold=bold
        if size: r.font.size=Pt(size)
        if color: r.font.color.rgb=RGBColor.from_string(color)
        r._element.get_or_add_rPr(); rf=r._element.rPr.find(qn('w:rFonts'))
        if rf is None: rf=OxmlElement('w:rFonts'); r._element.rPr.insert(0,rf)
        rf.set(qn('w:eastAsia'),east)
    if space_after is not None: p.paragraph_format.space_after=Pt(space_after)
    if br: p.add_run().add_break(WD_BREAK.PAGE)
    return p
for _ in range(5): para_before()
para_before('东营奥星石化 · 物料平衡与日效益',16,False,'1F7A8C',space_after=18)
para_before('物料平衡与日效益',28,True,'1F3864',space_after=6)
para_before('解决方案思路(可行性论证)',20,True,'1F3864',space_after=30)
tag=summary.split(':',1)[-1] if ':' in summary else summary
para_before(tag,12,False,'30404F',space_after=60)
para_before('版本:V0.1 讨论稿',12,False,'30404F',space_after=4)
para_before('日期:2026 年 11 月',12,False,'30404F',space_after=4)
para_before('说明:本文为现场调研前的方案思路,具体内容以调研结果为准;推测内容一律标"待确认"',10.5,False,'5A6570',space_after=0,br=True)
para_before('目  录',16,True,'1F3864',space_after=12)
# TOC field with static placeholder result
p=first.insert_paragraph_before(); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=0.5
def fld(p,kind,instr=None):
    r=OxmlElement('w:r')
    if kind=='instr':
        t=OxmlElement('w:instrText'); t.set(qn('xml:space'),'preserve'); t.text=instr; r.append(t)
    else:
        f=OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'),kind)
        if kind=='begin': f.set(qn('w:dirty'),'true')
        r.append(f)
    p._element.append(r)
fld(p,'begin'); fld(p,'instr',' TOC \\o "1-2" \\h \\z \\u '); fld(p,'separate')
last=p
for lv,txt in heads:
    q=first.insert_paragraph_before(); 
    q.paragraph_format.left_indent=Pt(0 if lv=='Heading 1' else 18); q.paragraph_format.space_after=Pt(0); q.paragraph_format.line_spacing=1.15
    q.paragraph_format.tab_stops.add_tab_stop(Twips(9366-60), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    r=q.add_run(txt+'\t'+str(PM.get(txt,''))); r.bold=(lv=='Heading 1'); last=q
fld(last,'end')
para_before('(如修改正文后页码有变化,请在 Word 中右键目录选择"更新域"。)',9,False,'8A94A0',space_after=0,br=True)
# make first content heading not double page break (cover/TOC already ends with break)
d.paragraphs[[i for i,x in enumerate(d.paragraphs) if x.style.name=='Heading 1'][0]].paragraph_format.page_break_before=False
# ---------- header / footer (default; first page blank)
hp=s.header.paragraphs[0]; hp.text=''; r=hp.add_run(HEADER); r.font.size=Pt(9); r.font.color.rgb=RGBColor.from_string('5A6570'); hp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
pPr=hp._element.get_or_add_pPr(); b=OxmlElement('w:pBdr'); bt=OxmlElement('w:bottom')
for k,v in (('w:val','single'),('w:sz','4'),('w:space','1'),('w:color','A7BFDD')): bt.set(qn(k),v)
b.append(bt); pPr.append(b)
fp=s.footer.paragraphs[0]; fp.text=''; fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
fp.add_run('第 ')
fb=OxmlElement('w:r'); x=OxmlElement('w:fldChar'); x.set(qn('w:fldCharType'),'begin'); fb.append(x); fp._element.append(fb)
fi=OxmlElement('w:r'); it=OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text=' PAGE '; fi.append(it); fp._element.append(fi)
fs=OxmlElement('w:r'); x=OxmlElement('w:fldChar'); x.set(qn('w:fldCharType'),'separate'); fs.append(x); fp._element.append(fs)
fp.add_run('1')
fe=OxmlElement('w:r'); x=OxmlElement('w:fldChar'); x.set(qn('w:fldCharType'),'end'); fe.append(x); fp._element.append(fe)
fp.add_run(' 页')
for r in fp.runs: r.font.size=Pt(9)
s.first_page_header.paragraphs[0].text=''; s.first_page_footer.paragraphs[0].text=''
# ---------- settings: update fields on open
setting=d.settings.element

# core props
cp=d.core_properties; cp.title='奥星石化物料平衡与日效益解决方案思路 V0.1'; cp.author=''; cp.last_modified_by=''; cp.subject='解决方案思路(可行性论证)'; cp.comments=''; cp.revision=1
d.save(dst); print('saved',dst, 'heads',len(heads))
