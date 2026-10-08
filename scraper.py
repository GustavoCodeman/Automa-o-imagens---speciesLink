#!/usr/bin/env python3
"""Coleta padrões de barcode do catálogo público de imagens do CRIA."""
import argparse, hashlib, re, sqlite3, time
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urljoin, urlparse
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup
from openpyxl import Workbook

BASE = "https://images.cria.org.br/"
CATEGORIES = ("herbaria", "micro", "musea", "photo-library")

def get(url, data=None):
    req = Request(url, data=data, headers={"User-Agent":"cria-barcode-scraper/1.0"})
    with urlopen(req, timeout=45) as r: return r.read().decode("utf-8", "replace")

def barcode(name):
    return Path(name).stem.split("_", 1)[0] or None

def pattern(value):
    parts = re.findall(r"\d+|\D+", value)
    label = " + ".join(f"{len(x)} dígitos" if x.isdigit() else x for x in parts)
    regex = "^" + "".join(r"\d{%d}" % len(x) if x.isdigit() else re.escape(x) for x in parts) + "$"
    return label, regex

def collections():
    soup = BeautifulSoup(get(urljoin(BASE,"osd-tree")), "html.parser")
    result = []
    for a in soup.select("a[href]"):
        q = parse_qs(urlparse(a["href"]).query); path = q.get("path",[""])[0].strip("/")
        if not path or "/" not in path or path.split("/",1)[0] not in CATEGORIES: continue
        try: count = int(q.get("count",[a.get("title","0")])[0])
        except ValueError: count = 0
        result.append((path, path.split("/",1)[0], count))
    return sorted(set(result))

def offsets(path, count):
    html = get(urljoin(BASE,"osd-head")+"?"+urlencode({"path":path,"count":count}))
    values = [int(x["value"]) for x in BeautifulSoup(html,"html.parser").select("select[name=offset] option[value]")]
    return values or list(range(0,count,200))

def names(path, offset):
    data = urlencode({"path":path,"offset":offset,"orderby":"imagecode asc","thumb_size":"medium"}).encode()
    soup = BeautifulSoup(get(urljoin(BASE,"osd-thumbs"),data),"html.parser")
    prefix = path+"/"; out=[]
    for img in soup.select("img[title]"):
        item = img["title"].split(maxsplit=1)[0]
        if item.startswith(prefix): out.append(item.rsplit("/",1)[-1])
    return out

def main():
    p=argparse.ArgumentParser(); p.add_argument("--state",default="data/progress.sqlite"); p.add_argument("--output",default="data/padroes_barcode_cria.xlsx"); p.add_argument("--delay",type=float,default=.25); a=p.parse_args()
    Path(a.state).parent.mkdir(parents=True,exist_ok=True); db=sqlite3.connect(a.state)
    db.execute("create table if not exists image(path text,name text,barcode text,pattern text,regex text, primary key(path,name))")
    db.execute("create table if not exists page(path text,offset integer, primary key(path,offset))")
    for path, category, expected in collections():
        for offset in offsets(path,expected):
            if db.execute("select 1 from page where path=? and offset=?",(path,offset)).fetchone(): continue
            current=names(path,offset)
            if not current: raise RuntimeError(f"Página vazia: {path} offset {offset}")
            for name in set(current):
                code=barcode(name); display,rx=pattern(code) if code else (None,None)
                db.execute("insert or ignore into image values(?,?,?,?,?)",(path,name,code,display,rx))
            db.execute("insert into page values(?,?)",(path,offset)); db.commit(); time.sleep(a.delay)
    wb=Workbook(); ws=wb.active; ws.title="Padrões por coleção"; ws.append(["Categoria","Coleção","Padrão","Exemplo real","Qtd. imagens","Qtd. barcodes únicos","Expressão regular"])
    for row in db.execute("select substr(path,1,instr(path,'/')-1),path,pattern,min(barcode),count(*),count(distinct barcode),regex from image group by path,pattern,regex order by path,pattern"): ws.append(row)
    summary=wb.create_sheet("Barcodes identificados"); summary.append(["Coleção","Barcode","Padrão","Qtd. imagens","Exemplo de arquivo"])
    for row in db.execute("select path,barcode,pattern,count(*),min(name) from image group by path,barcode,pattern order by path,barcode"): summary.append(row)
    for sh in wb:
        sh.freeze_panes="A2"; sh.auto_filter.ref=sh.dimensions
        for col in sh.columns: sh.column_dimensions[col[0].column_letter].width=min(50,max(12,max(len(str(x.value or "")) for x in col)+2))
    Path(a.output).parent.mkdir(parents=True,exist_ok=True); wb.save(a.output); print(a.output)
if __name__ == "__main__": main()
