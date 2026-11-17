import subprocess,json,re,docx
n=int(re.search(r'Pages:\s+(\d+)',subprocess.run(['pdfinfo','out.pdf'],capture_output=True,text=True).stdout).group(1))
pages=[subprocess.run(['pdftotext','-layout','-f',str(i),'-l',str(i),'out.pdf','-'],capture_output=True,text=True).stdout for i in range(1,n+1)]
d=docx.Document('out.docx')
heads=[p.text.split('\t')[0] for p in d.paragraphs if p.style.name in('Heading 1','Heading 2')]
toc_end=max(i for i,pg in enumerate(pages[:5]) if '目' in pg and ('录' in pg or '更新域' in pg) or '更新域' in pg)
m={}; start=toc_end+1
for h in heads:
    key=re.sub(r'\s','',h)[:14]
    for i in range(start,n):
        if key in re.sub(r'\s','',pages[i]): m[h]=i+1; start=i; break
old=json.load(open('pagemap.json'))
print('pages',n,'toc_end',toc_end+1,'changed', m!=old, len(m))
json.dump(m,open('pagemap.json','w'),ensure_ascii=False)
